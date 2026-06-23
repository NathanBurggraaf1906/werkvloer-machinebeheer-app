import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  images: {
    remotePatterns: [
      {
        hostname: "1906makersvancharcuterie.sharepoint.com",
        protocol: "https",
      },
    ],
  },
};

export default nextConfig;
