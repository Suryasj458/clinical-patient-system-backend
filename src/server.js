import dotenv from "dotenv";
import app from "./app.js";
import { pool } from "./config/db.js";

dotenv.config();
const port = process.env.PORT || 5000;

async function start() {
  try {
    await pool.query("SELECT 1");
    app.listen(port, () => console.log(`API running at http://localhost:${port}`));
  } catch (error) {
    console.error("Could not connect to PostgreSQL:", error.message);
    process.exit(1);
  }
}
start();
