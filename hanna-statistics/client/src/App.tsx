import { useEffect, useState } from 'react';

function App() {
  const [health, setHealth] = useState<string>('checking...');

  useEffect(() => {
    fetch('/api/health')
      .then((res) => res.json())
      .then((data) => setHealth(data.status ?? 'ok'))
      .catch(() => setHealth('backend offline'));
  }, []);

  return (
    <div className="min-h-screen flex items-center justify-center bg-[hsl(var(--background))]">
      <div className="text-center space-y-6">
        <h1 className="text-4xl font-bold text-[#1B3A6B]">
          Hanna Statistics
        </h1>
        <p className="text-lg text-[hsl(var(--muted-foreground))]">
          Development environment ready
        </p>
        <div className="inline-flex items-center gap-2 rounded-lg border px-4 py-2 text-sm">
          <span
            className={`h-2 w-2 rounded-full ${
              health === 'ok' ? 'bg-green-500' : 'bg-yellow-500'
            }`}
          />
          Backend: <span className="font-mono">{health}</span>
        </div>
      </div>
    </div>
  );
}

export default App;
