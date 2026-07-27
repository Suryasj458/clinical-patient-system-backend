import { query } from "../config/db.js";

function safeJson(value, fallback) {
  if (value === undefined || value === null) {
    return fallback;
  }

  if (typeof value === "string") {
    try {
      return JSON.parse(value);
    } catch {
      return fallback;
    }
  }

  return value;
}

function toNullableNumber(value) {
  if (value === "" || value === null || value === undefined) {
    return null;
  }

  const parsed = Number(value);

  return Number.isFinite(parsed) ? parsed : null;
}

function toNullableInteger(value) {
  const parsed = toNullableNumber(value);

  return parsed === null ? null : Math.trunc(parsed);
}

function toNullableText(value) {
  const normalized = String(value ?? "").trim();

  return normalized || null;
}

function toNullableDate(value) {
  const normalized = String(value ?? "").trim();

  return normalized || null;
}

const normalizePatient = (body = {}) => {
  const fields = safeJson(body.fields, {});
  const labs = safeJson(body.labs, {});
  const treatments = safeJson(body.treatments, []);

  const height = toNullableNumber(fields.height);
  const weight = toNullableNumber(fields.weight);

  const calculatedBmi =
    height && weight
      ? Number((weight / (height / 100) ** 2).toFixed(1))
      : null;

  return {
    patientName: String(fields.patientName ?? "").trim(),
    ipNo: String(fields.ipNo ?? "").trim(),
    age: toNullableInteger(fields.age),
    sex: toNullableText(fields.sex),
    admissionDate: toNullableDate(fields.admissionDate),
    assessmentDate: toNullableDate(fields.assessmentDate),
    ward: toNullableText(fields.ward),
    bedNo: toNullableText(fields.bedNo),
    heightCm: height,
    weightKg: weight,
    bmi: calculatedBmi,
    consultant: toNullableText(fields.consultant),

    clinicalFields:
      fields && typeof fields === "object" && !Array.isArray(fields)
        ? fields
        : {},

    labs:
      labs && typeof labs === "object" && !Array.isArray(labs)
        ? labs
        : {},

    treatments: Array.isArray(treatments) ? treatments : [],
  };
};

function formatDate(value) {
  if (!value) {
    return "";
  }

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }

  return String(value).slice(0, 10);
}

const mapRow = (row) => ({
  id: row.id,

  fields: {
    ...(row.clinical_fields || {}),

    patientName: row.patient_name ?? "",
    ipNo: row.ip_no ?? "",
    age: row.age ?? "",
    sex: row.sex ?? "",
    admissionDate: formatDate(row.admission_date),
    assessmentDate: formatDate(row.assessment_date),
    ward: row.ward ?? "",
    bedNo: row.bed_no ?? "",
    height: row.height_cm ?? "",
    weight: row.weight_kg ?? "",
    bmi: row.bmi ?? "",
    consultant: row.consultant ?? "",
  },

  labs: row.laboratory_results || {},
  treatments: Array.isArray(row.treatments) ? row.treatments : [],

  createdAt: row.created_at,
  updatedAt: row.updated_at,
});

export async function listPatients(req, res, next) {
  try {
    const search = String(req.query.search ?? "").trim();

    const requestedLimit = Number(req.query.limit);
    const requestedOffset = Number(req.query.offset);

    const limit = Number.isFinite(requestedLimit)
      ? Math.min(Math.max(requestedLimit, 1), 100)
      : 25;

    const offset = Number.isFinite(requestedOffset)
      ? Math.max(requestedOffset, 0)
      : 0;

    const pattern = `%${search}%`;

    const recordsQuery = query(
      `
        SELECT *
        FROM patients
        WHERE (
          $1 = ''
          OR patient_name ILIKE $2
          OR ip_no ILIKE $2
        )
        ORDER BY updated_at DESC
        LIMIT $3
        OFFSET $4
      `,
      [search, pattern, limit, offset],
    );

    const countQuery = query(
      `
        SELECT COUNT(*)::integer AS total
        FROM patients
        WHERE (
          $1 = ''
          OR patient_name ILIKE $2
          OR ip_no ILIKE $2
        )
      `,
      [search, pattern],
    );

    const [records, count] = await Promise.all([
      recordsQuery,
      countQuery,
    ]);

    return res.status(200).json({
      data: records.rows.map(mapRow),
      total: count.rows[0]?.total ?? 0,
      limit,
      offset,
    });
  } catch (error) {
    next(error);
  }
}

export async function getPatient(req, res, next) {
  try {
    const result = await query(
      `
        SELECT *
        FROM patients
        WHERE id = $1
      `,
      [req.params.id],
    );

    if (!result.rowCount) {
      return res.status(404).json({
        message: "Patient record not found.",
      });
    }

    return res.status(200).json(mapRow(result.rows[0]));
  } catch (error) {
    next(error);
  }
}

export async function createPatient(req, res, next) {
  try {
    const patient = normalizePatient(req.body);

    if (!patient.patientName || !patient.ipNo) {
      return res.status(400).json({
        message: "Patient Name and IP No. / UHID are required.",
      });
    }

    const result = await query(
      `
        INSERT INTO patients (
          patient_name,
          ip_no,
          age,
          sex,
          admission_date,
          assessment_date,
          ward,
          bed_no,
          height_cm,
          weight_kg,
          bmi,
          consultant,
          clinical_fields,
          laboratory_results,
          treatments
        )
        VALUES (
          $1,
          $2,
          $3,
          $4,
          $5,
          $6,
          $7,
          $8,
          $9,
          $10,
          $11,
          $12,
          $13::jsonb,
          $14::jsonb,
          $15::jsonb
        )
        RETURNING *
      `,
      [
        patient.patientName,
        patient.ipNo,
        patient.age,
        patient.sex,
        patient.admissionDate,
        patient.assessmentDate,
        patient.ward,
        patient.bedNo,
        patient.heightCm,
        patient.weightKg,
        patient.bmi,
        patient.consultant,

        JSON.stringify(patient.clinicalFields),
        JSON.stringify(patient.labs),
        JSON.stringify(patient.treatments),
      ],
    );

    return res.status(201).json(mapRow(result.rows[0]));
  } catch (error) {
    next(error);
  }
}

export async function updatePatient(req, res, next) {
  try {
    const patient = normalizePatient(req.body);

    if (!patient.patientName || !patient.ipNo) {
      return res.status(400).json({
        message: "Patient Name and IP No. / UHID are required.",
      });
    }

    const result = await query(
      `
        UPDATE patients
        SET
          patient_name = $1,
          ip_no = $2,
          age = $3,
          sex = $4,
          admission_date = $5,
          assessment_date = $6,
          ward = $7,
          bed_no = $8,
          height_cm = $9,
          weight_kg = $10,
          bmi = $11,
          consultant = $12,
          clinical_fields = $13::jsonb,
          laboratory_results = $14::jsonb,
          treatments = $15::jsonb
        WHERE id = $16
        RETURNING *
      `,
      [
        patient.patientName,
        patient.ipNo,
        patient.age,
        patient.sex,
        patient.admissionDate,
        patient.assessmentDate,
        patient.ward,
        patient.bedNo,
        patient.heightCm,
        patient.weightKg,
        patient.bmi,
        patient.consultant,

        JSON.stringify(patient.clinicalFields),
        JSON.stringify(patient.labs),
        JSON.stringify(patient.treatments),

        req.params.id,
      ],
    );

    if (!result.rowCount) {
      return res.status(404).json({
        message: "Patient record not found.",
      });
    }

    return res.status(200).json(mapRow(result.rows[0]));
  } catch (error) {
    next(error);
  }
}

export async function deletePatient(req, res, next) {
  try {
    const result = await query(
      `
        DELETE FROM patients
        WHERE id = $1
        RETURNING id
      `,
      [req.params.id],
    );

    if (!result.rowCount) {
      return res.status(404).json({
        message: "Patient record not found.",
      });
    }

    return res.status(204).send();
  } catch (error) {
    next(error);
  }
}