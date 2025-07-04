import { Router } from "express";
import microsoftGraphService from "../services/microsoftGraphService";

const router = Router();

router.post("/microsoft/callback", async (req, res) => {
  try {
    const userId = "auth0UserId"; // Replace with actual user ID from request context etc

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

    if (!tokenData || !tokenData.access_token) {
      res.status(400).json({
        error: "Failed to exchange authorization code for tokens",
      });
      return;
    }

    // ToDo: Save refresh token to Key Vault

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
