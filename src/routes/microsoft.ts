import { Router } from "express";
import microsoftGraphService from "../services/microsoftGraphService";

const router = Router();

router.get("/folders", async (req, res) => {
  try {
    const userId = "auth0UserId"; // Replace with actual user ID from request context etc
    const accessToken = await microsoftGraphService.getValidAccessToken(userId);

    const oneDriveFolders = await microsoftGraphService.getFolders(accessToken);

    const sharedWithMeFolders =
      await microsoftGraphService.getSharedWithMeFolders(accessToken);

    const sharePointSites = await microsoftGraphService.getSharePointSites(
      accessToken
    );

    const folders = [
      ...sharedWithMeFolders.map((folder: any) => ({
        ...folder,
        type: "shared",
      })),
      ...oneDriveFolders.map((folder: any) => ({
        ...folder,
        type: "onedrive",
      })),
      ...sharePointSites.map((site: any) => ({
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
      res.status(404).json({
        error: "Microsoft account not connected",
      });
      return;
    }

    res.status(500).json({
      error: "Failed to fetch folders",
    });
  }
});

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
