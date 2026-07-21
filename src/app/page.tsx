export default function Home() {
  return (
    <main className="min-h-screen bg-[#f7f8f3] text-[#20231d]">
      <section className="mx-auto flex min-h-screen w-full max-w-6xl flex-col px-6 py-8 sm:px-10">
        <nav className="flex items-center justify-between border-b border-[#dfe2d3] pb-5">
          <div>
            <p className="text-sm font-semibold uppercase tracking-[0.18em] text-[#667060]">
              AI Feedback
            </p>
            <h1 className="mt-2 text-2xl font-semibold">Workspace Starter</h1>
          </div>
          <div className="rounded-md border border-[#cad0bd] bg-white px-3 py-2 text-sm font-medium">
            Next.js + Supabase
          </div>
        </nav>

        <div className="grid flex-1 gap-8 py-10 lg:grid-cols-[1.1fr_0.9fr] lg:items-center">
          <div className="max-w-2xl">
            <p className="text-sm font-semibold uppercase tracking-[0.18em] text-[#6f7a66]">
              Ready for GitHub, Vercel, and Supabase
            </p>
            <h2 className="mt-4 text-4xl font-semibold leading-tight sm:text-6xl">
              Build the feedback system from this local folder.
            </h2>
            <p className="mt-6 max-w-xl text-lg leading-8 text-[#555d50]">
              This project is already scaffolded with Next.js, TypeScript,
              Tailwind CSS, and a Supabase client. Add your Supabase keys, push
              to GitHub, then import the repository in Vercel.
            </p>
          </div>

          <div className="rounded-lg border border-[#d9decb] bg-white p-5 shadow-sm">
            <div className="flex items-center justify-between border-b border-[#edf0e6] pb-4">
              <h3 className="text-lg font-semibold">Setup checklist</h3>
              <span className="rounded-md bg-[#e8f1dc] px-2 py-1 text-xs font-semibold text-[#526641]">
                Local ready
              </span>
            </div>
            <ol className="mt-5 space-y-4 text-sm text-[#4f574b]">
              <li className="rounded-md bg-[#f7f8f3] p-4">
                <strong className="block text-[#20231d]">1. Supabase</strong>
                Copy `.env.example` to `.env.local` and add project URL plus
                anon key.
              </li>
              <li className="rounded-md bg-[#f7f8f3] p-4">
                <strong className="block text-[#20231d]">2. GitHub</strong>
                Create a repository and push this folder as the main branch.
              </li>
              <li className="rounded-md bg-[#f7f8f3] p-4">
                <strong className="block text-[#20231d]">3. Vercel</strong>
                Import the GitHub repository and add the same environment
                variables.
              </li>
            </ol>
          </div>
        </div>
      </section>
    </main>
  );
}
