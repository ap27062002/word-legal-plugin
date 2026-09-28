import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  async headers() {
    return [
      {
        // Office loads the task pane in an iframe; relax the default
        // frame-ancestors restriction for Office domains during MVP.
        source: "/taskpane",
        headers: [
          {
            key: "Content-Security-Policy",
            value: "frame-ancestors 'self' https://*.officeapps.live.com https://*.office.com;",
          },
        ],
      },
    ];
  },
};

export default nextConfig;
