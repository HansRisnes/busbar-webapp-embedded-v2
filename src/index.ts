import { getMarketSnapshot } from "./services/marketDataService.js";

async function main(): Promise<void> {
  const snapshot = await getMarketSnapshot();
  console.log(JSON.stringify(snapshot, null, 2));
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  console.error(`Fatal error: ${message}`);
  process.exitCode = 1;
});
