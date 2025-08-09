import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";
import { Message } from "@microsoft/microsoft-graph-types";
import axios from "axios";
import dotenv from "dotenv";
import { AzureOpenAI } from "openai";
import { app } from "@azure/functions";

const env = process.env.NODE_ENV || "development";
if (env === "development") {
    dotenv.config({ path: ".env.local" });
}
else {
    dotenv.config();
}


const {
    CLIENT_ID,
    CLIENT_SECRET,
    TENANT_ID,
    AZURE_OPENAI_API_KEY,
    AZURE_OPENAI_API_ENDPOINT,
    AZURE_OPENAI_API_VERSION,
    AZURE_OPENAI_DEPLOYMENT_NAME,
    SUMMARY_TARGET_EMAIL,
    TARGET_USER_ID,
    SENDER_EMAIL_USER_ID
} = process.env;

async function getAccessToken(): Promise<string> {
    const credential = new ClientSecretCredential(TENANT_ID!, CLIENT_ID!, CLIENT_SECRET!);
    const token = await credential.getToken("https://graph.microsoft.com/.default");
    return token?.token!;
}

async function getGraphClient(): Promise<Client> {
    const token = await getAccessToken();
    return Client.init({
        authProvider: (done) => done(null, token),
    });
}

async function fetchUnreadEmails(graphClient: Client): Promise<Message[]> {
    const response = await graphClient
        .api(`users/${TARGET_USER_ID}/mailFolders/Inbox/messages`)
        .filter('isRead eq false')
        .get();

    return response.value;
}

async function sendSummaryEmail(summary: string): Promise<void> {
    if (!SUMMARY_TARGET_EMAIL) {
        console.warn("No summary target email configured.");
        return;
    }

    // Here you would implement the logic to send the summary email.
    // This is a placeholder function.
    console.log(`Sending summary email to ${SUMMARY_TARGET_EMAIL}:\n${summary}`);

    // use MS Graph API to send the email

    // Note: Implement the email sending logic using the Graph API.
    const graphClient = await getGraphClient();
    await graphClient.api(`/users/${SENDER_EMAIL_USER_ID}/sendMail`).post({
        message: {
            subject: "Email Summary",
            body: {
                contentType: "Text",
                content: summary
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: SUMMARY_TARGET_EMAIL
                    }
                }
            ]
        }
    });
}


interface EmailClassification {
    category: string;
    importance: string;
    suggestedAction: string;
    summary: string;
}

async function classifyWithOpenAI(subject: string, body: string): Promise<EmailClassification | null> {
    const prompt = `
You are an email assistant. Classify this email and suggest what to do.

Subject: ${subject}
Body: ${body}

Respond with JSON:
{
  "category": "work|personal|spam|newsletter|urgent|meeting|invoice|support",
  "importance": "low|medium|high|critical",
  "suggestedAction": "reply|forward|archive|delete|schedule_meeting|follow_up|no_action",
  "summary": "Brief 1-2 sentence summary of the email content"
}
`;

    try {
        const azureOpenAIOptions = {
            apiKey: AZURE_OPENAI_API_KEY,
            endpoint: AZURE_OPENAI_API_ENDPOINT,
            apiVersion: AZURE_OPENAI_API_VERSION,
            deployment: AZURE_OPENAI_DEPLOYMENT_NAME,
        }

        const azureOpenAIClient = new AzureOpenAI(azureOpenAIOptions);
        const response = await azureOpenAIClient.chat.completions.create({
            model: AZURE_OPENAI_DEPLOYMENT_NAME!,
            messages: [{ role: "user", content: prompt }],
            temperature: 0.2,
        });

        const content = response.choices[0].message.content;
        if (!content) {
            console.warn("No content received from OpenAI");
            return null;
        }

        return JSON.parse(content) as EmailClassification;
    } catch (err) {
        console.warn("Failed to classify email:", err);
        return null;
    }
}

async function processEmails(emails: Message[]): Promise<string> {
    const classifications: Array<{
        email: Message;
        classification: EmailClassification | null;
    }> = [];

    console.log(`Processing ${emails.length} unread emails...`);

    for (const email of emails) {
        const subject = email.subject || "No Subject";
        const body = email.body?.content || "No Content";

        console.log(`Classifying: ${subject}`);
        const classification = await classifyWithOpenAI(subject, body);

        classifications.push({
            email,
            classification
        });

        // Optional: Mark email as read after processing
        // await markEmailAsRead(graphClient, email.id!);
    }

    // Generate summary report
    return generateSummaryReport(classifications);
}

function generateSummaryReport(classifications: Array<{
    email: Message;
    classification: EmailClassification | null;
}>): string {
    const summary = [`Email Summary Report - ${new Date().toLocaleDateString()}\n`];

    const categoryCounts: Record<string, number> = {};
    const highPriorityEmails: Array<{ email: Message, classification: EmailClassification }> = [];

    classifications.forEach(({ email, classification }) => {
        if (classification) {
            categoryCounts[classification.category] = (categoryCounts[classification.category] || 0) + 1;

            if (classification.importance === 'high' || classification.importance === 'critical') {
                highPriorityEmails.push({ email, classification });
            }
        }
    });

    // Category breakdown
    summary.push("📊 Category Breakdown:");
    Object.entries(categoryCounts).forEach(([category, count]) => {
        summary.push(`  ${category}: ${count} emails`);
    });

    // High priority emails
    if (highPriorityEmails.length > 0) {
        summary.push("\n🚨 High Priority Emails:");
        highPriorityEmails.forEach(({ email, classification }) => {
            summary.push(`  • ${email.subject} [${classification.importance}]`);
            summary.push(`    From: ${email.from?.emailAddress?.address}`);
            summary.push(`    Action: ${classification.suggestedAction}`);
            summary.push(`    Summary: ${classification.summary}\n`);
        });
    }

    // All email summaries
    summary.push("\n📧 All Email Summaries:");
    classifications.forEach(({ email, classification }) => {
        if (classification) {
            summary.push(`\n• ${email.subject || 'No Subject'}`);
            summary.push(`  From: ${email.from?.emailAddress?.address || 'Unknown'}`);
            summary.push(`  Category: ${classification.category} | Priority: ${classification.importance}`);
            summary.push(`  Suggested Action: ${classification.suggestedAction}`);
            summary.push(`  Summary: ${classification.summary}`);
        }
    });

    return summary.join('\n');
}

async function markEmailAsRead(graphClient: Client, emailId: string): Promise<void> {
    try {
        await graphClient
            .api(`/users/${TARGET_USER_ID}/messages/${emailId}`)
            .patch({
                isRead: true
            });
    } catch (error) {
        console.warn(`Failed to mark email ${emailId} as read:`, error);
    }
}

// Updated main execution
export async function main(): Promise<void> {
    try {
        console.log("Starting email scanner...");

        // Validate environment variables
        if (!CLIENT_ID || !CLIENT_SECRET || !TENANT_ID) {
            console.error("Missing required Azure AD credentials");
            process.exit(1);
        }

        if (!AZURE_OPENAI_API_KEY || !AZURE_OPENAI_API_ENDPOINT) {
            console.error("Missing required Azure OpenAI credentials");
            process.exit(1);
        }

        const client = await getGraphClient();
        const emails = await fetchUnreadEmails(client);

        if (emails.length === 0) {
            console.log("No unread emails found.");
            return;
        }

        const summaryReport = await processEmails(emails);

        // Send summary email if configured
        if (SUMMARY_TARGET_EMAIL) {
            await sendSummaryEmail(summaryReport);
            console.log(`Summary sent to ${SUMMARY_TARGET_EMAIL}`);
        } else {
            console.log("Summary Report:");
            console.log(summaryReport);
        }

    } catch (error) {
        console.error("Error running email scanner:", error);
        process.exit(1);
    }
}

app.timer('timetrigger', {
    schedule: '0 1 * * 6',
    handler: async () => {
        await main();
    }
});