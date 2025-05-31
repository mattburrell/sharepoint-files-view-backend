import axios from "axios";

class MicrosoftGraphService {
  clientId: string | undefined;
  clientSecret: string | undefined;
  baseURL: string;

  constructor() {
    this.clientId = process.env.MICROSOFT_CLIENT_ID;
    this.clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
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
      const tokenUrl =
        "https://login.microsoftonline.com/common/oauth2/v2.0/token";

      const params = new URLSearchParams({
        client_id: this.clientId,
        client_secret: this.clientSecret,
        scope: "Files.Read Sites.Read.All offline_access",
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
      const tokenUrl =
        "https://login.microsoftonline.com/common/oauth2/v2.0/token";

      const params = new URLSearchParams({
        client_id: this.clientId,
        client_secret: this.clientSecret,
        scope: "Files.Read Sites.Read.All offline_access",
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

  async getValidAccessToken(auth0UserId) {
    return "";
    // const userToken = await UserToken.findOne({ auth0UserId });

    // if (!userToken) {
    //   throw new Error("No Microsoft tokens found for user");
    // }

    // const now = new Date();
    // const expiresAt = new Date(userToken.microsoftTokens.expiresAt);

    // // If token expires within 5 minutes, refresh it
    // if (expiresAt <= new Date(now.getTime() + 5 * 60 * 1000)) {
    //   const refreshedTokens = await this.refreshAccessToken(
    //     userToken.microsoftTokens.refreshToken
    //   );

    //   userToken.microsoftTokens.accessToken = refreshedTokens.access_token;
    //   if (refreshedTokens.refresh_token) {
    //     userToken.microsoftTokens.refreshToken = refreshedTokens.refresh_token;
    //   }
    //   userToken.microsoftTokens.expiresAt = new Date(
    //     now.getTime() + refreshedTokens.expires_in * 1000
    //   );

    //   await userToken.save();

    //   return refreshedTokens.access_token;
    // }

    // return userToken.microsoftTokens.accessToken;
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
