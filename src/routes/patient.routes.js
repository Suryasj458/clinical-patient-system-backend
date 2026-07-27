import { Router } from "express";
import {
  createPatient,
  deletePatient,
  getPatient,
  listPatients,
  updatePatient,
} from "../controllers/patient.controller.js";

const router = Router();
router.get("/", listPatients);
router.get("/:id", getPatient);
router.post("/", createPatient);
router.put("/:id", updatePatient);
router.delete("/:id", deletePatient);
export default router;
