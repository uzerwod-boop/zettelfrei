import { Assistant } from "./components/Assistant";
import { Belege, Dokumente, Pruefen, Verlauf } from "./components/Panels";
import { Settings } from "./components/Settings";
import { QualityTest } from "./components/QualityTest";
import { UpdateBar } from "./components/UpdateBar";
import { Feedback } from "./components/Feedback";
import { Kpis, Rail, Sheet, Toast, TopBar } from "./components/Workspace";
import { useStore } from "./store";

export default function App() {
  const s = useStore();
  return (
    <div className="app">
      <Rail />
      <main className="center"><TopBar /><Kpis /><Sheet /></main>
      <aside className="panel" aria-label="Seitenbereich">
        <div hidden={s.pane !== "cmd"} className="panewrap"><Assistant /></div>
        <div hidden={s.pane !== "pdf"} className="panewrap"><Belege /></div>
        <Dokumente active={s.pane === "docs"} />
        {s.pane === "check" && <Pruefen />}
        {s.pane === "hist" && <Verlauf />}
      </aside>
      <Settings /><QualityTest /><UpdateBar /><Feedback /><Toast />
    </div>);
}
