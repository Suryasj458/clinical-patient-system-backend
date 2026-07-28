import ExcelJS from "exceljs";
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

function formatExcelHeading(value = "") {
  return String(value)
    .replace(/([A-Z])/g, " $1")
    .replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .replace(/\b\w/g, (character) => character.toUpperCase())
    .trim();
}

function normalizeExcelValue(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof Date) {
    return value;
  }

  if (Array.isArray(value)) {
    return value
      .map((item) =>
        item && typeof item === "object"
          ? JSON.stringify(item)
          : String(item),
      )
      .join(", ");
  }

  if (typeof value === "object") {
    return JSON.stringify(value);
  }

  return value;
}

function collectObjectKeys(rows, propertyName) {
  const keys = new Set();

  rows.forEach((row) => {
    const value = row[propertyName];

    if (
      value &&
      typeof value === "object" &&
      !Array.isArray(value)
    ) {
      Object.keys(value).forEach((key) => keys.add(key));
    }
  });

  return Array.from(keys);
}

function collectTreatmentKeys(rows) {
  const keys = new Set();

  rows.forEach((row) => {
    const treatments = Array.isArray(row.treatments)
      ? row.treatments
      : [];

    treatments.forEach((treatment) => {
      if (
        treatment &&
        typeof treatment === "object" &&
        !Array.isArray(treatment)
      ) {
        Object.keys(treatment).forEach((key) => keys.add(key));
      }
    });
  });

  return Array.from(keys);
}

function styleWorksheet(worksheet) {
  worksheet.views = [
    {
      state: "frozen",
      ySplit: 1,
    },
  ];

  if (worksheet.columnCount > 0) {
    worksheet.autoFilter = {
      from: {
        row: 1,
        column: 1,
      },
      to: {
        row: 1,
        column: worksheet.columnCount,
      },
    };
  }

  const headerRow = worksheet.getRow(1);

  headerRow.height = 32;

  headerRow.font = {
    bold: true,
    color: {
      argb: "FFFFFFFF",
    },
  };

  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
      argb: "FF1E3A5F",
    },
  };

  headerRow.alignment = {
    vertical: "middle",
    horizontal: "center",
    wrapText: true,
  };

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.alignment = {
        vertical: "top",
        wrapText: true,
      };
    }

    row.eachCell((cell) => {
      cell.border = {
        top: {
          style: "thin",
          color: { argb: "FFD9E1EA" },
        },
        left: {
          style: "thin",
          color: { argb: "FFD9E1EA" },
        },
        bottom: {
          style: "thin",
          color: { argb: "FFD9E1EA" },
        },
        right: {
          style: "thin",
          color: { argb: "FFD9E1EA" },
        },
      };
    });
  });
}

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


export async function exportPatientsExcel(req, res, next) {
  try {
    const search = String(req.query.search ?? "").trim();
    const pattern = `%${search}%`;

    const result = await query(
      `
        SELECT
          id,
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
          treatments,
          created_at,
          updated_at
        FROM patients
        WHERE (
          $1 = ''
          OR patient_name ILIKE $2
          OR ip_no ILIKE $2
        )
        ORDER BY updated_at DESC
      `,
      [search, pattern],
    );

    const records = result.rows;

    if (!records.length) {
      return res.status(404).json({
        message: "No patient records found for Excel export.",
      });
    }

    const clinicalFieldKeys = collectObjectKeys(
      records,
      "clinical_fields",
    );

    const laboratoryKeys = collectObjectKeys(
      records,
      "laboratory_results",
    );

    const treatmentKeys = collectTreatmentKeys(records);

    const workbook = new ExcelJS.Workbook();

    workbook.creator = "Clinical Patient Entry System";
    workbook.lastModifiedBy = "Clinical Patient Entry System";
    workbook.created = new Date();
    workbook.modified = new Date();

    /*
     * Sheet 1: Patient details
     */
    const patientWorksheet = workbook.addWorksheet("Patients");

    patientWorksheet.columns = [
      {
        header: "Record ID",
        key: "id",
        width: 38,
      },
      {
        header: "Patient Name",
        key: "patientName",
        width: 26,
      },
      {
        header: "IP No. / UHID",
        key: "ipNo",
        width: 20,
      },
      {
        header: "Age",
        key: "age",
        width: 10,
      },
      {
        header: "Sex",
        key: "sex",
        width: 14,
      },
      {
        header: "Admission Date",
        key: "admissionDate",
        width: 18,
      },
      {
        header: "Assessment Date",
        key: "assessmentDate",
        width: 18,
      },
      {
        header: "Ward",
        key: "ward",
        width: 18,
      },
      {
        header: "Bed No.",
        key: "bedNo",
        width: 14,
      },
      {
        header: "Height (cm)",
        key: "heightCm",
        width: 14,
      },
      {
        header: "Weight (kg)",
        key: "weightKg",
        width: 14,
      },
      {
        header: "BMI",
        key: "bmi",
        width: 12,
      },
      {
        header: "Consultant",
        key: "consultant",
        width: 24,
      },

      ...clinicalFieldKeys.map((key) => ({
        header: formatExcelHeading(key),
        key: `clinical_${key}`,
        width: 24,
      })),

      {
        header: "Created At",
        key: "createdAt",
        width: 22,
      },
      {
        header: "Updated At",
        key: "updatedAt",
        width: 22,
      },
    ];

    records.forEach((record) => {
      const clinicalFields =
        record.clinical_fields &&
        typeof record.clinical_fields === "object"
          ? record.clinical_fields
          : {};

      const excelRow = {
        id: record.id,
        patientName: record.patient_name ?? "",
        ipNo: record.ip_no ?? "",
        age: record.age ?? "",
        sex: record.sex ?? "",
        admissionDate: record.admission_date || "",
        assessmentDate: record.assessment_date || "",
        ward: record.ward ?? "",
        bedNo: record.bed_no ?? "",
        heightCm: record.height_cm ?? "",
        weightKg: record.weight_kg ?? "",
        bmi: record.bmi ?? "",
        consultant: record.consultant ?? "",
        createdAt: record.created_at
          ? new Date(record.created_at)
          : "",
        updatedAt: record.updated_at
          ? new Date(record.updated_at)
          : "",
      };

      clinicalFieldKeys.forEach((key) => {
        excelRow[`clinical_${key}`] = normalizeExcelValue(
          clinicalFields[key],
        );
      });

      patientWorksheet.addRow(excelRow);
    });

    patientWorksheet.getColumn("admissionDate").numFmt =
      "dd-mm-yyyy";

    patientWorksheet.getColumn("assessmentDate").numFmt =
      "dd-mm-yyyy";

    patientWorksheet.getColumn("createdAt").numFmt =
      "dd-mm-yyyy hh:mm";

    patientWorksheet.getColumn("updatedAt").numFmt =
      "dd-mm-yyyy hh:mm";

    styleWorksheet(patientWorksheet);

    /*
     * Sheet 2: Laboratory results
     */
    const laboratoryWorksheet = workbook.addWorksheet(
      "Laboratory Results",
    );

    laboratoryWorksheet.columns = [
      {
        header: "Record ID",
        key: "recordId",
        width: 38,
      },
      {
        header: "Patient Name",
        key: "patientName",
        width: 26,
      },
      {
        header: "IP No. / UHID",
        key: "ipNo",
        width: 20,
      },

      ...laboratoryKeys.map((key) => ({
        header: formatExcelHeading(key),
        key: `lab_${key}`,
        width: 22,
      })),
    ];

    records.forEach((record) => {
      const laboratoryResults =
        record.laboratory_results &&
        typeof record.laboratory_results === "object"
          ? record.laboratory_results
          : {};

      const excelRow = {
        recordId: record.id,
        patientName: record.patient_name ?? "",
        ipNo: record.ip_no ?? "",
      };

      laboratoryKeys.forEach((key) => {
        excelRow[`lab_${key}`] = normalizeExcelValue(
          laboratoryResults[key],
        );
      });

      laboratoryWorksheet.addRow(excelRow);
    });

    styleWorksheet(laboratoryWorksheet);

    /*
     * Sheet 3: Treatments
     */
    const treatmentWorksheet = workbook.addWorksheet("Treatments");

    treatmentWorksheet.columns = [
      {
        header: "Record ID",
        key: "recordId",
        width: 38,
      },
      {
        header: "Patient Name",
        key: "patientName",
        width: 26,
      },
      {
        header: "IP No. / UHID",
        key: "ipNo",
        width: 20,
      },
      {
        header: "Treatment Number",
        key: "treatmentNumber",
        width: 18,
      },

      ...treatmentKeys.map((key) => ({
        header: formatExcelHeading(key),
        key: `treatment_${key}`,
        width: 24,
      })),
    ];

    records.forEach((record) => {
      const treatments =
        Array.isArray(record.treatments) &&
        record.treatments.length
          ? record.treatments
          : [{}];

      treatments.forEach((treatment, index) => {
        const excelRow = {
          recordId: record.id,
          patientName: record.patient_name ?? "",
          ipNo: record.ip_no ?? "",
          treatmentNumber: index + 1,
        };

        treatmentKeys.forEach((key) => {
          excelRow[`treatment_${key}`] =
            normalizeExcelValue(treatment?.[key]);
        });

        treatmentWorksheet.addRow(excelRow);
      });
    });

    styleWorksheet(treatmentWorksheet);

    const date = new Date().toISOString().slice(0, 10);
    const fileName = search
      ? `clinical-patients-${search.replace(
          /[^a-zA-Z0-9-_]/g,
          "-",
        )}-${date}.xlsx`
      : `clinical-patients-${date}.xlsx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${fileName}"`,
    );

    res.setHeader(
      "Access-Control-Expose-Headers",
      "Content-Disposition",
    );

    await workbook.xlsx.write(res);

    return res.end();
  } catch (error) {
    next(error);
  }
}