/** @type {import('next').NextConfig} */
const nextConfig = {
  typescript: {
    ignoreBuildErrors: true,
  },
  images: {
    unoptimized: true,
  },
  experimental: {
    serverActions: {
      bodySizeLimit: '50mb',
    },
  },
  // Increase body parser limit for API route handlers
  api: {
    bodyParser: {
      sizeLimit: '50mb',
    },
  },
}

export default nextConfig
