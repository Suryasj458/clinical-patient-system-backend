CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

CREATE TABLE IF NOT EXISTS patients (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  patient_name VARCHAR(180) NOT NULL,
  ip_no VARCHAR(100) NOT NULL UNIQUE,
  age INTEGER CHECK (age IS NULL OR age >= 0),
  sex VARCHAR(20),
  admission_date DATE,
  assessment_date DATE,
  ward VARCHAR(150),
  bed_no VARCHAR(50),
  height_cm NUMERIC(7,2),
  weight_kg NUMERIC(7,2),
  bmi NUMERIC(6,2),
  consultant VARCHAR(180),
  clinical_fields JSONB NOT NULL DEFAULT '{}'::jsonb,
  laboratory_results JSONB NOT NULL DEFAULT '{}'::jsonb,
  treatments JSONB NOT NULL DEFAULT '[]'::jsonb,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_patients_name_lower
  ON patients (LOWER(patient_name));

CREATE INDEX IF NOT EXISTS idx_patients_assessment_date
  ON patients (assessment_date DESC);

CREATE OR REPLACE FUNCTION set_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS patients_set_updated_at ON patients;
CREATE TRIGGER patients_set_updated_at
BEFORE UPDATE ON patients
FOR EACH ROW EXECUTE FUNCTION set_updated_at();
