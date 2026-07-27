import express from "express";
import cors from "cors";
import helmet from "helmet";
import morgan from "morgan";
import patientRoutes from "./routes/patient.routes.js";
import { errorHandler, notFound } from "./middleware/errorHandler.js";

const app = express();
app.use(helmet());
app.use(cors({ origin: process.env.CLIENT_URL || "http://localhost:5173" }));
app.use(express.json({ limit: "5mb" }));
app.use(morgan("dev"));
app.get("/api/health", (req, res) => res.json({ status: "ok" }));
app.use("/api/patients", patientRoutes);
app.use(notFound);
app.use(errorHandler);
export default app;
