import { createServer, build } from 'vite';
import { spawn } from 'child_process';
import electronPath from 'electron';

async function startDev() {
  // 1. Build electron main and preload scripts
  await build({ configFile: './vite.electron.config.ts' });

  // 2. Start Vite dev server
  const server = await createServer({
    configFile: './vite.config.ts',
  });
  await server.listen();
  const address = server.httpServer?.address();
  const port = typeof address === 'object' && address ? address.port : 5173;
  console.log(`Vite dev server running at http://localhost:${port}`);

  // 3. Launch Electron
  const env = {
    ...process.env,
    VITE_DEV_SERVER_URL: `http://localhost:${port}`,
  };

  const electronProcess = spawn(electronPath, ['.'], {
    stdio: 'inherit',
    env,
  });

  electronProcess.on('close', () => {
    server.close();
    process.exit();
  });
}

startDev().catch((err) => {
  console.error('Error starting dev environment:', err);
  process.exit(1);
});
