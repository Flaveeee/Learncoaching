import { createFileRoute } from '@tanstack/react-router'
import { useState, useEffect, useRef, useCallback } from 'react'

export const Route = createFileRoute('/')({
  component: LerncoachingApp,
})

// ── Data ───────────────────────────────────────────────────────────────

const PHASES = [
  {
    id: 'einstieg',
    label: 'Einstieg',
    title: 'Ankommen & Einstieg',
    titleEm: '',
    desc: 'Beziehung aufbauen und Gesprächsrahmen setzen',
    questions: [
      { text: 'Wie geht es dir heute? Was beschäftigt dich gerade?' },
      { text: 'Was hast du seit unserem letzten Gespräch erlebt – in der Schule oder zu Hause?' },
      { text: 'Gibt es etwas, das du mir besonders mitteilen möchtest, bevor wir beginnen?' },
      { text: 'Wie war deine Woche bisher? Was war gut, was eher schwierig?' },
    ],
  },
  {
    id: 'analyse',
    label: 'Analyse',
    title: 'Situation & Analyse',
    titleEm: '',
    desc: 'Lernstand, Stärken und Herausforderungen erkunden',
    questions: [
      { text: 'In welchen Fächern oder Bereichen fühlst du dich <strong>sicher und stark</strong>? Was klappt gut?' },
      { text: 'Wo merkst du, dass es dir <strong>schwerer fällt</strong>? Was genau ist dabei schwierig?' },
      { text: 'Wie gehst du normalerweise beim Lernen vor? Beschreibe mir deinen Alltag.' },
      { text: 'Wann lernst du am besten – zu welcher Tageszeit, in welcher Umgebung?' },
      { text: 'Was hat dir bisher beim Lernen geholfen? Was hat nicht funktioniert?' },
    ],
  },
  {
    id: 'ziele',
    label: 'Ziele',
    title: 'Ziele & Wünsche',
    titleEm: '',
    desc: 'Konkrete, erreichbare Ziele gemeinsam entwickeln',
    questions: [
      { text: 'Was möchtest du in den nächsten Wochen <strong>erreichen oder verbessern</strong>?' },
      { text: 'Stell dir vor, du schaust in drei Monaten zurück: Was soll sich verändert haben?' },
      { text: 'Was ist dir beim Lernen wichtig? Was motiviert dich?' },
      { text: 'Gibt es ein Ziel, das du dir selbst gesetzt hast – auch ohne dass jemand es von dir verlangt?' },
      { text: 'Wie kann ich dich dabei unterstützen, dieses Ziel zu erreichen?' },
    ],
  },
  {
    id: 'strategien',
    label: 'Strategien',
    title: 'Strategien & Planung',
    titleEm: '',
    desc: 'Wirksame Lernwege entdecken und erproben',
    questions: [
      { text: 'Was könntest du konkret ausprobieren, um dein Ziel zu erreichen?' },
      { text: 'Kennst du Techniken wie Mindmaps, Lernkarten oder lautes Erklären? Was davon könnte für dich passen?' },
      { text: 'Wann und wie lange willst du in der nächsten Woche üben oder lernen?' },
      { text: 'Was könnte dich ablenken oder stören? Wie gehst du damit um?' },
      { text: 'Wen kannst du um Hilfe bitten, wenn du nicht weiterkommst?' },
    ],
  },
  {
    id: 'reflexion',
    label: 'Reflexion',
    title: 'Reflexion & Abschluss',
    titleEm: '',
    desc: 'Das Gespräch verankern und nächste Schritte vereinbaren',
    questions: [
      { text: 'Was nimmst du aus unserem Gespräch heute mit?' },
      { text: 'Was war für dich am <strong>wertvollsten oder überraschendsten</strong>?' },
      { text: 'Was ist dein konkreter nächster Schritt – was machst du bis zu unserem nächsten Treffen?' },
      { text: 'Auf einer Skala von 1 bis 10: Wie zuversichtlich bist du, dass du diesen Schritt umsetzen wirst?' },
      { text: 'Gibt es noch etwas, das du sagen möchtest, bevor wir abschließen?' },
    ],
  },
]

// ── Timer hook ───────────────────────────────────────────────────────────

function useTimer() {
  const [elapsed, setElapsed] = useState(0)
  const [running, setRunning] = useState(false)
  const intervalRef = useRef<ReturnType<typeof setInterval> | null>(null)

  const start = useCallback(() => {
    if (!running) {
      setRunning(true)
      intervalRef.current = setInterval(() => setElapsed(e => e + 1), 1000)
    } else {
      setRunning(false)
      if (intervalRef.current) clearInterval(intervalRef.current)
    }
  }, [running])

  const reset = useCallback(() => {
    setRunning(false)
    setElapsed(0)
    if (intervalRef.current) clearInterval(intervalRef.current)
  }, [])

  useEffect(() => () => { if (intervalRef.current) clearInterval(intervalRef.current) }, [])

  const mm = String(Math.floor(elapsed / 60)).padStart(2, '0')
  const ss = String(elapsed % 60).padStart(2, '0')

  return { display: `${mm}:${ss}`, running, start, reset, elapsed }
}

// ── Main component ────────────────────���──────────────────────────────────

function LerncoachingApp() {
  const [students, setStudents] = useState<string[]>([])
  const [activeStudent, setActiveStudent] = useState<string | null>(null)
  const [newName, setNewName] = useState('')
  const [phaseIdx, setPhaseIdx] = useState(0)
  const [donePhases, setDonePhases] = useState<Set<number>>(new Set())
  const [highlighted, setHighlighted] = useState<Set<string>>(new Set())
  const [notes, setNotes] = useState<Record<string, string>>({})
  const [showHelp, setShowHelp] = useState(false)
  const [animKey, setAnimKey] = useState(0)
  const timer = useTimer()

  const phase = PHASES[phaseIdx]
  const noteKey = `${activeStudent ?? '__none'}__${phase.id}`
  const noteVal = notes[noteKey] ?? ''

  function addStudent() {
    const trimmed = newName.trim()
    if (!trimmed || students.includes(trimmed)) return
    setStudents(s => [...s, trimmed])
    setActiveStudent(trimmed)
    setNewName('')
  }

  function removeStudent(name: string) {
    setStudents(s => s.filter(n => n !== name))
    if (activeStudent === name) setActiveStudent(null)
  }

  function goPhase(idx: number) {
    setDonePhases(d => new Set([...d, phaseIdx]))
    setPhaseIdx(idx)
    setAnimKey(k => k + 1)
  }

  function toggleHighlight(qId: string) {
    setHighlighted(h => {
      const next = new Set(h)
      next.has(qId) ? next.delete(qId) : next.add(qId)
      return next
    })
  }

  function updateNote(val: string) {
    setNotes(n => ({ ...n, [noteKey]: val }))
  }

  const sanitizeFilename = (value: string) => {
    const cleaned = value
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '')

    return cleaned || 'kind'
  }

  const handleDownloadStudentNotes = () => {
    if (!activeStudent) return

    const lines = PHASES.map((phaseItem) => {
      const phaseNote = notes[`${activeStudent}__${phaseItem.id}`]?.trim() || 'Keine Notizen vorhanden.'
      return `## ${phaseItem.label}: ${phaseItem.title}\n${phaseNote}\n`
    }).join('\n---\n\n')

    const content = [
      `Lerncoaching – ${activeStudent}`,
      `Datum: ${new Date().toLocaleDateString('de-DE')}`,
      `Gesprächsdauer: ${timer.display}`,
      '',
      lines,
    ].join('\n')

    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = `lerncoaching_${sanitizeFilename(activeStudent)}_${new Date().toISOString().slice(0, 10)}.txt`
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    URL.revokeObjectURL(url)
  }

  const highlightCount = [...highlighted].filter(k => k.startsWith(phase.id)).length

  return (
    <div className="app-shell">
      {/* Header */}
      <header className="site-header">
        <div className="site-logo">
          Lerncoaching <span>· Gesprächsführer</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
          {activeStudent && (
            <span className="tag">{activeStudent}</span>
          )}
          <button className="help-btn" onClick={() => setShowHelp(true)} title="Hilfe">?</button>
        </div>
      </header>

      {/* Main */}
      <main className="main-grid">

        {/* Left sidebar */}
        <aside className="sidebar-left">

          {/* Students */}
          <div className="card">
            <div className="card-header">
              <span className="card-label">Schülerinnen & Schüler</span>
            </div>
            <div className="card-body">
              {students.length > 0 && (
                <div className="student-list">
                  {students.map(name => (
                    <div
                      key={name}
                      className={`student-chip ${activeStudent === name ? 'active' : ''}`}
                      onClick={() => setActiveStudent(name)}
                    >
                      <span>{name}</span>
                      <button
                        className="remove-btn"
                        onClick={e => { e.stopPropagation(); removeStudent(name) }}
                        title="Entfernen"
                      >×</button>
                    </div>
                  ))}
                </div>
              )}
              {students.length === 0 && (
                <p style={{ fontSize: '0.8125rem', color: 'var(--ink-muted)', marginBottom: '0.75rem', fontStyle: 'italic' }}>
                  Noch keine Schüler:innen hinzugefügt.
                </p>
              )}
              <div className="student-add-form">
                <input
                  className="input-base"
                  placeholder="Name eingeben …"
                  value={newName}
                  onChange={e => setNewName(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && addStudent()}
                />
                <button className="btn-icon" onClick={addStudent} title="Hinzufügen">+</button>
              </div>
            </div>
          </div>

          {/* Timer */}
          <div className="card">
            <div className="card-header">
              <span className="card-label">Gesprächszeit</span>
            </div>
            <div className="timer-display">{timer.display}</div>
            <div className="timer-controls">
              <button
                className={`btn-sm ${timer.running ? '' : 'primary'}`}
                onClick={timer.start}
              >
                {timer.running ? 'Pause' : 'Start'}
              </button>
              <button className="btn-sm danger" onClick={timer.reset}>Zurücksetzen</button>
            </div>
            <div style={{ height: '0.75rem' }} />
          </div>

          {/* Phase progress */}
          <div className="card">
            <div className="card-header">
              <span className="card-label">Gesprächsphasen</span>
            </div>
            <div className="phase-progress">
              {PHASES.map((p, i) => (
                <div
                  key={p.id}
                  className={`phase-step ${i === phaseIdx ? 'active' : ''} ${donePhases.has(i) ? 'done' : ''}`}
                  onClick={() => goPhase(i)}
                >
                  <div className="step-dot" />
                  <span className="step-label">{p.label}</span>
                  {donePhases.has(i) && i !== phaseIdx && (
                    <svg style={{ marginLeft: 'auto', width: 12, height: 12, color: 'var(--sage)' }} viewBox="0 0 12 12" fill="none">
                      <path d="M2 6l3 3 5-5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                    </svg>
                  )}
                </div>
              ))}
            </div>
          </div>

        </aside>

        {/* Center: phase content */}
        <section className="content-main">
          {!activeStudent ? (
            <div className="card phase-card">
              <div className="empty-state">
                <div className="empty-icon">
                  <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="var(--forest)" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/>
                    <circle cx="9" cy="7" r="4"/>
                    <path d="M23 21v-2a4 4 0 0 0-3-3.87"/>
                    <path d="M16 3.13a4 4 0 0 1 0 7.75"/>
                  </svg>
                </div>
                <h3>Schüler:in auswählen</h3>
                <p>Füge eine Schülerin oder einen Schüler hinzu und wähle sie oder ihn aus, um das Gespräch zu beginnen.</p>
              </div>
            </div>
          ) : (
            <div className="card phase-card phase-enter" key={animKey}>
              <div className="phase-card-header">
                <div className="phase-number">Phase {phaseIdx + 1} von {PHASES.length}</div>
                <div className="phase-title">{phase.title}{phase.titleEm && <em> {phase.titleEm}</em>}</div>
                <div className="phase-desc">{phase.desc}</div>
              </div>

              <div className="questions-area">
                {phase.questions.map((q, qi) => {
                  const qId = `${phase.id}__${qi}`
                  return (
                    <div
                      key={qId}
                      className={`question-item ${highlighted.has(qId) ? 'highlighted' : ''}`}
                      onClick={() => toggleHighlight(qId)}
                    >
                      <span className="q-number">{qi + 1}</span>
                      <span
                        className="q-text"
                        dangerouslySetInnerHTML={{ __html: q.text }}
                      />
                      <button
                        className="q-highlight-btn"
                        onClick={e => { e.stopPropagation(); toggleHighlight(qId) }}
                      >
                        {highlighted.has(qId) ? 'markiert' : 'markieren'}
                      </button>
                    </div>
                  )
                })}
              </div>

              <div className="phase-nav">
                <button
                  className="btn-nav"
                  onClick={() => goPhase(phaseIdx - 1)}
                  disabled={phaseIdx === 0}
                >
                  ← Zurück
                </button>

                <div className="phase-dots">
                  {PHASES.map((_, i) => (
                    <div
                      key={i}
                      className={`phase-dot ${i === phaseIdx ? 'active' : ''} ${donePhases.has(i) && i !== phaseIdx ? 'done' : ''}`}
                      onClick={() => goPhase(i)}
                    />
                  ))}
                </div>

                <button
                  className={`btn-nav next ${phaseIdx < PHASES.length - 1 ? 'primary-nav' : ''}`}
                  onClick={() => phaseIdx < PHASES.length - 1 ? goPhase(phaseIdx + 1) : undefined}
                  disabled={phaseIdx === PHASES.length - 1}
                >
                  Weiter →
                </button>
              </div>
            </div>
          )}
        </section>

        {/* Right sidebar */}
        <aside className="sidebar-right">

          {/* Notes */}
          <div className="card" style={{ flex: 1 }}>
            <div className="card-header">
              <span className="card-label">Notizen</span>
              {activeStudent && <span style={{ fontSize: '0.75rem', color: 'var(--ink-muted)', marginLeft: 'auto' }}>{phase.label}</span>}
            </div>
            <textarea
              className="notes-textarea"
              placeholder={activeStudent ? 'Beobachtungen, Zitate, Ideen …' : 'Wähle zuerst eine:n Schüler:in aus …'}
              value={noteVal}
              onChange={e => updateNote(e.target.value)}
              disabled={!activeStudent}
            />
          </div>

          {/* Session summary */}
          {activeStudent && (
            <div className="card">
              <div className="card-header">
                <span className="card-label">Sitzungsübersicht</span>
              </div>
              <div className="card-body" style={{ padding: '1rem 1.25rem' }}>
                <div className="summary-card">
                  <div className="summary-row">
                    <span>Schüler:in</span>
                    <span style={{ fontWeight: 500 }}>{activeStudent}</span>
                  </div>
                  <div className="summary-row">
                    <span>Aktuelle Phase</span>
                    <span>{phase.label}</span>
                  </div>
                  <div className="summary-row">
                    <span>Gesprächsdauer</span>
                    <span style={{ fontFamily: 'Fraunces, serif' }}>{timer.display}</span>
                  </div>
                  <div className="summary-row">
                    <span>Markierte Fragen</span>
                    <span>{highlightCount} / {phase.questions.length}</span>
                  </div>
                  <div className="summary-row">
                    <span>Phasen abgeschlossen</span>
                    <span>{donePhases.size} / {PHASES.length}</span>
                  </div>
                </div>

                <button
                  className="btn-sm primary"
                  onClick={handleDownloadStudentNotes}
                  style={{ width: '100%', marginTop: '1rem' }}
                >
                  Speichern
                </button>
              </div>
            </div>
          )}

        </aside>
      </main>

      {/* Help modal */}
      {showHelp && (
        <div className="help-overlay" onClick={() => setShowHelp(false)}>
          <div className="help-modal" onClick={e => e.stopPropagation()}>
            <h2>So nutzt du den Gesprächsführer</h2>
            <p>Dieses Tool begleitet dich strukturiert durch ein Lerncoaching-Gespräch.</p>
            <ul>
              <li>Füge eine Schülerin oder einen Schüler in der linken Spalte hinzu und wähle sie oder ihn aus.</li>
              <li>Starte den Timer, wenn das Gespräch beginnt.</li>
              <li>Klicke auf eine Frage, um sie hervorzuheben – so behältst du den Überblick.</li>
              <li>Notiere deine Beobachtungen in der rechten Spalte – pro Phase und Schüler:in separat.</li>
              <li>Navigiere mit «Zurück» und «Weiter» durch die fünf Phasen.</li>
            </ul>
            <p>Die Phasen sind: Einstieg → Analyse → Ziele → Strategien → Reflexion.</p>
            <button className="help-close" onClick={() => setShowHelp(false)}>Verstanden</button>
          </div>
        </div>
      )}
    </div>
  )
}
