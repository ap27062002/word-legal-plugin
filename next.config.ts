import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  async headers() {
    return [
      {
        // Office (desktop and web) frames more than just /taskpane during
        // validation and runtime handshakes, so this applies site-wide.
        source: "/:path*",
        headers: [
          {
            key: "Content-Security-Policy",
            value:
              "frame-ancestors 'self' https://appsforoffice.microsoft.com https://*.office.com https://*.officeapps.live.com https://*.cloud.microsoft https://*.sharepoint.com https://*.resources.office.net;",
          },
        ],
      },
    ];
  },
};

export default nextConfig;
