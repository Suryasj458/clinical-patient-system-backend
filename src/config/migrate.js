import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { pool } from "./db.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

try {
  const sqlPath = path.resolve(__dirname, "../../sql/schema.sql");
  const sql = await fs.readFile(sqlPath, "utf8");
  await pool.query(sql);
  console.log("Database migration completed.");
} catch (error) {
  console.error("Migration failed:", error);
  process.exitCode = 1;
} finally {
  await pool.end();
}
