import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";
import { Message } from "@microsoft/microsoft-graph-types";
import axios from "axios";
import dotenv from "dotenv";

dotenv.config();

const {
    CLIENT_ID,
    CLIENT_SECRET,
    TENANT_ID,
    AZURE_OPENAI_API_KEY,
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
        .api('//mailFolders/Inbox/messages')
        .filter('isRead eq false')
        .select('id,subject,bodyPreview,from')
        .get();

    return response.value;
}

async function classifyWithOpenAI(subject: string, body: string): Promise<string> {
    const prompt = `
You are an email assistant. Classify this email and suggest what to do.

Subject: ${subject}
Body: ${body}

Respond with JSON:
{
  "category": "...",
  "importance": "...",
  "suggestedAction": "...",
  "summary": "..."
}
`;

    const response = await axios.post("https://api.openai.com/v1/chat/completions", {
        model: "gpt-4",
        messages: [{ role: "user", content: prompt }],
        temperature: 0.2,
    }, {
        headers: {
            Authorization: `Bearer ${process.env.AZURE_OPENAI_API_KEY}`,
        },
    });

    const content = response.data.choices[0].message.content;
    try {
        return JSON.parse(content);
    } catch (err) {
        console.warn("Failed to parse OpenAI response:", content);
        return null;
    }
}

(async () => {
    const client = await getGraphClient();
    const emails = await fetchUnreadEmails(client);


    for (const email of emails) {
        const result = await classifyWithOpenAI(email.subject, email.bodyPreview);
        console.log(`\n📧 Subject: ${email.subject}`);
        console.log(`👤 From: ${email.from.emailAddress.address}`);
        console.log(`📋 Result:`, result);
    }
})();