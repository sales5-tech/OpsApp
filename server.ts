import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { google } from "googleapis";
import { GoogleAuth } from "google-auth-library";
import dotenv from "dotenv";
import admin from "firebase-admin";
import cryptoRandomString from "crypto-random-string";

dotenv.config();

// Initialize Firebase Admin
const serviceAccountJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
try {
  if (serviceAccountJson) {
    const credentials = JSON.parse(serviceAccountJson);
    if (!admin.apps.length) {
      admin.initializeApp({
        credential: admin.credential.cert(credentials)
      });
    }
  } else if (!admin.apps.length) {
    // Try default initialization if running in GCP environment
    admin.initializeApp();
  }
} catch (err) {
  console.error("Firebase Admin initialization error:", err);
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: '50mb' }));

  // Initialize db inside startServer to ensure admin is initialized
  const getDb = () => {
    if (admin.apps.length) return admin.firestore();
    return null;
  };

  // Google Sheets Integration
  const getSheetsClient = () => {
    const serviceAccountJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
    if (!serviceAccountJson) {
      throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is not defined");
    }
    const credentials = JSON.parse(serviceAccountJson);
    const auth = new GoogleAuth({
      credentials,
      scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });
    return google.sheets({ version: "v4", auth });
  };

  app.post("/api/sync-to-sheets", async (req, res) => {
    try {
      console.log("Starting Google Sheets Sync...");
      const { data, sheetName } = req.body;
      const spreadsheetId = process.env.GOOGLE_SHEET_ID;
      
      if (!spreadsheetId) {
        console.error("GOOGLE_SHEET_ID is missing");
        return res.status(400).json({ error: "GOOGLE_SHEET_ID is not defined in environment variables" });
      }

      console.log(`Syncing to Spreadsheet ID: ${spreadsheetId}, Sheet: ${sheetName}`);
      
      let sheets;
      try {
        sheets = getSheetsClient();
      } catch (authError: any) {
        console.error("Auth Error:", authError.message);
        return res.status(500).json({ error: `Authentication failed: ${authError.message}. Check GOOGLE_SERVICE_ACCOUNT_JSON.` });
      }
      
      // Check if sheet exists, if not create it
      console.log("Fetching spreadsheet metadata...");
      const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
      const sheetExists = spreadsheet.data.sheets?.some(s => s.properties?.title === sheetName);
      
      if (!sheetExists) {
        console.log(`Sheet '${sheetName}' not found. Creating it...`);
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [{ addSheet: { properties: { title: sheetName } } }]
          }
        });
      }

      // Clear and update sheet
      console.log(`Clearing range ${sheetName}!A1:Z10000...`);
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: `${sheetName}!A1:Z10000`,
      });

      console.log(`Updating values in ${sheetName}!A1...`);
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!A1`,
        valueInputOption: "RAW",
        requestBody: { values: data },
      });

      console.log("Sync completed successfully!");
      res.json({ success: true });
    } catch (error: any) {
      console.error("Google Sheets Sync Error:", error);
      res.status(500).json({ error: error.message || "Unknown error during sync" });
    }
  });

  // Shared Links Integration
  app.post("/api/create-share-link", async (req, res) => {
    const db = getDb();
    if (!db) return res.status(500).json({ error: "قاعدة البيانات غير متصلة - يرجى إضافة GOOGLE_SERVICE_ACCOUNT_JSON" });
    
    const { itemId, expiryDays } = req.body;
    const token = cryptoRandomString({ length: 32, type: 'alphanumeric' });
    
    const expiryDate = expiryDays 
      ? new Date(Date.now() + expiryDays * 24 * 60 * 60 * 1000).toISOString()
      : null;

    try {
      const linkRef = db.collection("shared_links").doc();
      await linkRef.set({
        token,
        itemId,
        expiryDate,
        viewCount: 0,
        active: true,
        createdAt: new Date().toISOString()
      });
      res.json({ success: true, token });
    } catch (error) {
      console.error("Create Share Link Error:", error);
      res.status(500).json({ error: "فشل إنشاء الرابط" });
    }
  });

  app.get("/api/shared-item/:token", async (req, res) => {
    const db = getDb();
    if (!db) return res.status(500).json({ error: "قاعدة البيانات غير متصلة - يرجى إضافة GOOGLE_SERVICE_ACCOUNT_JSON" });
    
    const { token } = req.params;
    
    try {
      const linksSnap = await db.collection("shared_links").where("token", "==", token).where("active", "==", true).limit(1).get();
      
      if (linksSnap.empty) {
        return res.status(404).json({ error: "الرابط غير صالح أو ملغي" });
      }

      const linkDoc = linksSnap.docs[0];
      const linkData = linkDoc.data();

      // Check Expiry
      if (linkData.expiryDate && new Date(linkData.expiryDate) < new Date()) {
        await linkDoc.ref.update({ active: false });
        return res.status(410).json({ error: "الرابط منتهي الصلاحية" });
      }

      // Update View Count
      await linkDoc.ref.update({ viewCount: admin.firestore.FieldValue.increment(1) });

      // Fetch Item
      const itemDoc = await db.collection("items").doc(linkData.itemId).get();
      
      if (!itemDoc.exists) {
        return res.status(404).json({ error: "البيانات الأصلية غير موجودة" });
      }

      res.json({ success: true, item: { id: itemDoc.id, ...itemDoc.data() } });
    } catch (error) {
      console.error("Fetch Shared Item Error:", error);
      res.status(500).json({ error: "فشل جلب البيانات" });
    }
  });

  app.post("/api/revoke-share-link", async (req, res) => {
    const db = getDb();
    if (!db) return res.status(500).json({ error: "قاعدة البيانات غير متصلة" });
    const { token } = req.body;
    try {
      const linksSnap = await db.collection("shared_links").where("token", "==", token).limit(1).get();
      if (!linksSnap.empty) {
        await linksSnap.docs[0].ref.update({ active: false });
      }
      res.json({ success: true });
    } catch (error) {
      res.status(500).json({ error: "فشل الحذف" });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
