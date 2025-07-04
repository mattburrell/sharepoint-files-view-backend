# SharePoint Graph API Backend

This backend service issues secure access tokens for SharePoint file access via the Microsoft Graph API. It serves as an authentication layer for the frontend, using delegated permissions and the authorization code OAuth flow.

## Getting Started

### Prerequisites

- Node.js (v20+)
- npm or yarn

### Installation

```bash
npm install
```

### Configuration

Copy `.env.example` to `.env` and update environment variables as needed.

### Running the Server

```bash
npm start
```

### Development

```bash
npm run dev
```

## Project Structure

```
backend/
├── src/
│   ├── models/
│   ├── routes/
│   └── services/
├── package.json
└── README.md
```

## License

[MIT](../LICENSE)
