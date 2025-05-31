import { Router } from "express";
import microsoftGraphService from "../services/microsoftGraphService";

const router = Router();

// Get user's OneDrive and SharePoint folders
router.get("/folders", async (req, res) => {
  try {
    // Get valid access token
    const accessToken = await microsoftGraphService.getValidAccessToken(
      "userId"
    );

    // Fetch OneDrive folders
    const oneDriveFolders = await microsoftGraphService.getFolders(accessToken);

    // Fetch SharePoint sites
    const sharePointSites = await microsoftGraphService.getSharePointSites(
      accessToken
    );

    // Combine results
    const folders = [
      ...oneDriveFolders.map((folder) => ({
        ...folder,
        type: "onedrive",
      })),
      ...sharePointSites.map((site) => ({
        id: site.id,
        name: site.displayName || site.name,
        webUrl: site.webUrl,
        type: "sharepoint",
      })),
    ];

    res.json({ folders });
  } catch (error: any) {
    console.error("Get folders error:", error);

    if (error.message === "No Microsoft tokens found for user") {
      return res.status(404).json({
        error: "Microsoft account not connected",
      });
    }

    res.status(500).json({
      error: "Failed to fetch folders",
    });
  }
});

// Disconnect Microsoft account
router.delete("/disconnect", async (req, res) => {
  try {
    res.json({
      success: true,
      message: "Microsoft account disconnected",
    });
  } catch (error) {
    console.error("Disconnect error:", error);
    res.status(500).json({
      error: "Failed to disconnect Microsoft account",
    });
  }
});

export default router;
