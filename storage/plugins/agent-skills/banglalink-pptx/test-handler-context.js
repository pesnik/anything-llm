/**
 * Test the handler with a simulated server execution context.
 * Mimics exactly what the aibitat does:
 *   fn.caller = "agent";
 *   const result = await fn.handler(args);
 * where `this` inside handler() === fn (the functionConfig).
 */

const path = require("path");
const fs = require("fs");

const handlerMod = require("./handler");

// ── Simulated aibitat instance ──────────────────────────────
const loggedMessages = [];
const introspectMessages = [];
const socketSentMessages = [];
const pendingOutputs = [];

const mockAibitat = {
  handlerProps: {
    log: (...args) => {
      loggedMessages.push(args.join(" "));
      console.log("[mock:handlerProps.log]", ...args);
    },
  },
  socket: {
    send: (type, content) => {
      socketSentMessages.push({ type, content });
      console.log(`[mock:socket.send] type="${type}"`, JSON.stringify(content));
    },
  },
  introspect: (msg) => {
    introspectMessages.push(msg);
    console.log(`[mock:introspect] "${msg}"`);
  },
  _pendingOutputs: pendingOutputs,
};

// ── functionConfig ────────────────────────────────────────────
const functionConfig = {
  super: mockAibitat,
  name: "banglalink-pptx",
  caller: "agent",
  logger: mockAibitat.handlerProps.log,
  config: { name: "Banglalink PPTX Generator", description: "..." },
  handler: handlerMod.handler,
};

// ── Test input ────────────────────────────────────────────────
const testArgs = {
  topic: "Test – Full Handler Context",
  sections: [
    {
      title: "Test Section 1",
      topic: "Testing",
      bulletPoints: ["Bullet A", "Bullet B", "Bullet C"],
    },
    {
      title: "Test Section 2",
      topic: "More Testing",
      isClosing: true,
      bulletPoints: ["Final point"],
    },
  ],
  audience: "Test Audience",
  outputFileName: "test-handler-context",
  theme: "dark",
  enableResearch: false,
};

async function run() {
  console.log("═".repeat(60));
  console.log("Handler Context Test – simulating aibitat.fn.handler(args)");
  console.log("═".repeat(60) + "\n");

  const result = await functionConfig.handler.call(functionConfig, testArgs);
  console.log(`\nHandler return value: "${result}"`);

  // ── Verify ─────────────────────────────────────────────
  console.log("\n─ Verifications ─────────────────────────────────");
  const isError = result.startsWith("Error:");
  console.log(`  Is error:         ${isError}`);
  if (isError) {
    console.error("  ✗ HANDLER RETURNED ERROR");
    process.exit(1);
  }

  // Check that saveGeneratedFile was called (file on disk)
  const createFilesLib = require(path.join(
    __dirname, "..", "..", "..", "..",
    "server", "utils", "agents", "aibitat",
    "plugins", "create-files", "lib.js"
  ));

  const storageRoot =
    process.env.STORAGE_DIR ||
    path.resolve(__dirname, "../../../../server/storage");
  const genDir = path.join(storageRoot, "generated-files");
  const recentFiles = fs
    .readdirSync(genDir)
    .filter((f) => f.endsWith(".pptx") && f.startsWith("pptx-"))
    .sort()
    .slice(-3);
  console.log(`  Recent files in ${genDir}: ${recentFiles.join(", ") || "NONE"}`);

  // Check socket messages
  console.log(`  Socket messages sent: ${socketSentMessages.length}`);
  for (const msg of socketSentMessages) {
    console.log(`    - type="${msg.type}" filename="${msg.content?.filename || "N/A"}"`);
  }
  const hasFileCard = socketSentMessages.some((m) => m.type === "fileDownloadCard");
  console.log(`  fileDownloadCard sent: ${hasFileCard ? "✓" : "✗ MISSING"}`);

  // Check introspect messages
  console.log(`  Introspect messages:  ${introspectMessages.length}`);

  // Check logger messages
  console.log(`  Logger messages:      ${loggedMessages.length}`);

  // Check pendingOutputs
  console.log(`  _pendingOutputs:      ${mockAibitat._pendingOutputs.length}`);

  const allOk = !isError && hasFileCard && mockAibitat._pendingOutputs.length > 0;
  console.log(`\n${allOk ? "✓ ALL CHECKS PASSED" : "✗ SOME CHECKS FAILED"}`);
}

run().catch((e) => {
  console.error("Unhandled error:", e.message);
  process.exit(1);
});
