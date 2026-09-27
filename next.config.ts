import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  eslint: {
    // Ignora gli errori di ESLint durante le build di produzione su Vercel
    ignoreDuringBuilds: true,
  },
};

export default nextConfig;