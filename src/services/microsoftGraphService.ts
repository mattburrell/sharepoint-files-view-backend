import axios from "axios";
import UserToken from "../models/userToken";

class MicrosoftGraphService {
  clientId: string | undefined;
  clientSecret: string | undefined;
  tenantId: string | undefined;
  baseURL: string;

  constructor() {
    this.clientId = process.env.MICROSOFT_CLIENT_ID;
    this.clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
    this.tenantId = process.env.MICROSOFT_TENANT_ID;
    this.baseURL = "https://graph.microsoft.com/v1.0";
  }

  async exchangeCodeForTokens(
    code: string,
    codeVerifier: string,
    redirectUri: string
  ) {
    if (!this.clientId || !this.clientSecret) {
      throw new Error("Microsoft client ID and secret are not configured");
    }

    try {
      const tokenUrl = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;

      const params = new URLSearchParams({
        client_id: this.clientId,
        client_secret: this.clientSecret,
        scope: "Files.Read offline_access",
        code: code,
        redirect_uri: redirectUri,
        grant_type: "authorization_code",
        code_verifier: codeVerifier,
      });

      const response = await axios.post(tokenUrl, params, {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
      });

      return response.data;
    } catch (error: any) {
      console.error(
        "Token exchange error:",
        error.response?.data || error.message
      );
      throw new Error("Failed to exchange authorization code for tokens");
    }
  }

  async refreshAccessToken(refreshToken: string) {
    if (!this.clientId || !this.clientSecret) {
      throw new Error("Microsoft client ID and secret are not configured");
    }
    if (!refreshToken) {
      throw new Error("Refresh token is required to refresh access token");
    }

    try {
      const tokenUrl = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;

      const params = new URLSearchParams({
        client_id: this.clientId,
        client_secret: this.clientSecret,
        scope: "Files.Read.All offline_access",
        refresh_token: refreshToken,
        grant_type: "refresh_token",
      });

      const response = await axios.post(tokenUrl, params, {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
      });

      return response.data;
    } catch (error: any) {
      console.error(
        "Token refresh error:",
        error.response?.data || error.message
      );
      throw new Error("Failed to refresh access token");
    }
  }

  async getValidAccessToken(auth0UserId: string) {
    // ToDo: Retrieve the user's refresh token from Key Vault and then get a valid access token
    return "access token";
  }

  async getUserDrives(accessToken: string) {
    try {
      const response = await axios.get(`${this.baseURL}/me/drives`, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      return response.data.value || [];
    } catch (error: any) {
      console.error("Get drives error:", error.response?.data || error.message);
      throw new Error("Failed to fetch user drives");
    }
  }

  async getSharedWithMeFolders(accessToken: string) {
    try {
      const response = await axios.get(
        `${this.baseURL}/me/drive/sharedWithMe?allowexternal=true`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      const allItems = response.data.value || [];
      const folders = allItems.filter((item: any) => item.folder != null);
      return folders;
    } catch (error: any) {
      console.error(
        "Get shared with me folders error:",
        error.response?.data || error.message
      );
      throw new Error("Failed to fetch shared with me folders");
    }
  }

  async getFolders(accessToken: string, driveId = null) {
    try {
      let url = `${this.baseURL}/me/drive/root/children`;
      if (driveId) {
        url = `${this.baseURL}/drives/${driveId}/root/children`;
      }

      const response = await axios.get(url, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
        params: {
          $filter: "folder ne null",
          $select: "id,name,webUrl,folder,parentReference",
        },
      });

      return response.data.value || [];
    } catch (error: any) {
      console.error(
        "Get folders error:",
        error.response?.data || error.message
      );
      throw new Error("Failed to fetch folders");
    }
  }

  async getFollowedSites(accessToken: string) {
    try {
      const response = await axios.get(`${this.baseURL}/me/followedSites`, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
        params: {
          $select: "id,name,webUrl,displayName",
        },
      });
      return response.data.value || [];
    } catch (error: any) {
      console.error(
        "Get followed sites error:",
        error.response?.data || error.message
      );
      throw new Error("Failed to fetch followed sites");
    }
  }

  async getSharePointSites(accessToken: string) {
    try {
      const response = await axios.get(`${this.baseURL}/sites?search=*`, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
        params: {
          $select: "id,name,webUrl,displayName",
        },
      });

      return response.data.value || [];
    } catch (error: any) {
      console.error(
        "Get SharePoint sites error:",
        error.response?.data || error.message
      );
      throw new Error("Failed to fetch SharePoint sites");
    }
  }
}

export default new MicrosoftGraphService();
