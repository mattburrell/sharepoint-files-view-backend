import express from "express";
import cors from "cors";
import oauthRoutes from "./routes/oauth";
import microsoftRoutes from "./routes/microsoft";

const app = express();
const port = 5000;

app.use(
  cors({
    origin: process.env.FRONTEND_URL || "http://localhost:5173",
    credentials: true,
  })
);

app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true }));

app.use("/api/oauth", oauthRoutes);
app.use("/api/microsoft", microsoftRoutes);

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
