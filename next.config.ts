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
            value: "frame-ancestors 'self' https://*.officeapps.live.com https://*.office.com https://*.live.com;",
          },
        ],
      },
    ];
  },
};

export default nextConfig;
