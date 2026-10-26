import { Assistant } from "./components/Assistant";
import { Settings } from "./components/Settings";
import { Kpis, Rail, Sheet, Toast, TopBar } from "./components/Workspace";
import { useStore } from "./store";

// Bereiche rechts, die noch aus der klassischen Oberfläche übernommen werden
function Soon({ title }: { title: string }) {
  return <div className="pane"><div className="panehead"><h2>{title}</h2></div>
    <p className="hint">Dieser Bereich kommt in der nächsten Version der neuen Oberfläche. Bis dahin funktioniert er in der <a href="../">klassischen Ansicht</a>.</p></div>;
}
export default function App() {
  const s = useStore();
  return (
    <div className="app">
      <Rail />
      <main className="center"><TopBar /><Kpis /><Sheet /></main>
      <aside className="panel" aria-label="Seitenbereich">
        {s.pane === "cmd" && <Assistant />}
        {s.pane === "pdf" && <Soon title="Belege übernehmen" />}
        {s.pane === "docs" && <Soon title="Dokumente" />}
        {s.pane === "check" && <Soon title="Was stimmt nicht?" />}
        {s.pane === "hist" && <Soon title="Verlauf" />}
      </aside>
      <Settings /><Toast />
    </div>);
}
