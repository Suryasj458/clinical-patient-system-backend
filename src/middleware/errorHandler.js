export function notFound(req, res) {
  res.status(404).json({ message: `Route not found: ${req.method} ${req.originalUrl}` });
}

export function errorHandler(error, req, res, next) {
  console.error(error);

  if (error.code === "23505") {
    return res.status(409).json({ message: "A patient with this IP No. / UHID already exists." });
  }

  const status = error.status || 500;
  return res.status(status).json({
    message: error.message || "Internal server error",
  });
}
