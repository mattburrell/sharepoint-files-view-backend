import { Router } from "express";
import microsoftGraphService from "../services/microsoftGraphService";

const router = Router();

router.post("/microsoft/callback", async (req, res) => {
  try {
    const { code, codeVerifier, redirectUri } = req.body;

    if (!code || !codeVerifier || !redirectUri) {
      res.status(400).json({
        error: "Missing required parameters",
      });
      return;
    }

    const tokenData = await microsoftGraphService.exchangeCodeForTokens(
      code,
      codeVerifier,
      redirectUri
    );

    const expiresAt = new Date(Date.now() + tokenData.expires_in * 1000);
    console.log("Token data received:", tokenData);
    console.log("Token expires at:", expiresAt);
    // Save tokens to the database (assuming you have a UserToken model)

    res.json({
      success: true,
      message: "Microsoft account connected successfully",
    });
  } catch (error) {
    console.error("OAuth callback error:", error);
    res.status(500).json({
      error: "Failed to process OAuth callback",
    });
  }
});

export default router;
