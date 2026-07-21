# AI Feedback

Next.js frontend starter for an AI feedback system using Supabase as the
backend, GitHub as the source repository, and Vercel as the deployment target.

## Local setup

The local Node.js runtime is installed in:

```powershell
C:\Users\chompoopan.jan\Tools\node-v24.18.0-win-x64
```

For this PowerShell session, add it to `Path` before running npm commands:

```powershell
$env:Path="$env:USERPROFILE\Tools\node-v24.18.0-win-x64;$env:Path"
```

Copy the environment example and fill in the real Supabase values:

```powershell
Copy-Item .env.example .env.local
```

```env
NEXT_PUBLIC_SUPABASE_URL=https://your-project-ref.supabase.co
NEXT_PUBLIC_SUPABASE_ANON_KEY=your-supabase-anon-key
```

## Getting Started

First, run the development server:

```bash
npm run dev
# or
yarn dev
# or
pnpm dev
# or
bun dev
```

Open [http://localhost:3000](http://localhost:3000) with your browser to see the result.

You can start editing the page by modifying `app/page.tsx`. The page auto-updates as you edit the file.

This project uses [`next/font`](https://nextjs.org/docs/app/building-your-application/optimizing/fonts) to automatically optimize and load [Geist](https://vercel.com/font), a new font family for Vercel.

## Deploy flow

1. Create a GitHub repository named `AI-Feedback` or `ai-feedback`.
2. Push this local folder to the repository.
3. Import the repository in Vercel.
4. Add the same Supabase environment variables in Vercel project settings.
5. Each push to the production branch will trigger a Vercel deployment.

## Learn More

To learn more about Next.js, take a look at the following resources:

- [Next.js Documentation](https://nextjs.org/docs) - learn about Next.js features and API.
- [Learn Next.js](https://nextjs.org/learn) - an interactive Next.js tutorial.

You can check out [the Next.js GitHub repository](https://github.com/vercel/next.js) - your feedback and contributions are welcome!

## Deploy on Vercel

The easiest way to deploy your Next.js app is to use the [Vercel Platform](https://vercel.com/new?utm_medium=default-template&filter=next.js&utm_source=create-next-app&utm_campaign=create-next-app-readme) from the creators of Next.js.

Check out our [Next.js deployment documentation](https://nextjs.org/docs/app/building-your-application/deploying) for more details.
