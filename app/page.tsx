export default function Home() {
  return (
    <main style={{ fontFamily: "sans-serif", padding: 24 }}>
      <h1>Legal Rulebook Copilot</h1>
      <p>
        This is the backend + task pane host for the Word add-in. The add-in
        UI itself lives at <code>/taskpane</code> and is loaded by Word via
        the manifest in <code>/manifest/manifest.xml</code>.
      </p>
    </main>
  );
}
