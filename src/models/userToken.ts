interface UserToken {
  auth0UserId: string;
  microsoftTokens: {
    accessToken: string;
    refreshToken: string;
    expiresAt: Date;
  };
  createdAt: Date;
  updatedAt: Date;
  lastUsedAt: Date;
  microsoftAccountId?: string; // Optional, in case the user has multiple accounts
  microsoftAccountName?: string; // Optional, for display purposes
  microsoftAccountEmail?: string; // Optional, for display purposes
  microsoftAccountPicture?: string; // Optional, for display purposes
  microsoftAccountType?: "onedrive" | "sharepoint"; // Optional, to distinguish account types
}

export default UserToken;
