import "dotenv/config";

import express from "express";
import cors from "cors";
import helmet from "helmet";
import morgan from "morgan";

import patientRoutes from "./routes/patient.routes.js";
import {
  errorHandler,
  notFound,
} from "./middleware/errorHandler.js";

const app = express();

const allowedOrigins = [
  "http://localhost:5173",
  "https://clinical-patient-system-frontend.vercel.app",
  process.env.CLIENT_URL,
]
  .filter(Boolean)
  .map((url) => url.replace(/\/$/, ""));

const corsOptions = {
  origin(origin, callback) {
    // Allows Postman, curl and direct browser requests
    if (!origin) {
      return callback(null, true);
    }

    const normalizedOrigin = origin.replace(/\/$/, "");

    if (allowedOrigins.includes(normalizedOrigin)) {
      return callback(null, true);
    }

    console.error("CORS blocked origin:", origin);
    console.log("Allowed origins:", allowedOrigins);

    return callback(new Error(`Origin ${origin} is not allowed by CORS`));
  },

  credentials: true,

  methods: [
    "GET",
    "POST",
    "PUT",
    "PATCH",
    "DELETE",
    "OPTIONS",
  ],

  allowedHeaders: [
    "Content-Type",
    "Authorization",
    "Accept",
  ],
};

app.use(cors(corsOptions));
app.use(helmet());
app.use(express.json({ limit: "5mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(morgan("dev"));

app.get("/api/health", (req, res) => {
  res.status(200).json({
    status: "ok",
    clientUrl: process.env.CLIENT_URL || null,
  });
});

app.use("/api/patients", patientRoutes);

app.use(notFound);
app.use(errorHandler);

export default app;