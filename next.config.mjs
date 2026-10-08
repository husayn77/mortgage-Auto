/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  experimental: {
    serverComponentsExternalPackages: ["exceljs", "sharp", "jszip"],
  },
};

export default nextConfig;
