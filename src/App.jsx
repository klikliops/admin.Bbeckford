import { useEffect, useId, useMemo, useRef, useState } from 'react';
import {
  ArrowRight,
  ArrowDownAZ,
  CalendarClock,
  Check,
  ChevronDown,
  Clock3,
  CreditCard,
  Eye,
  EyeOff,
  FileSpreadsheet,
  LayoutDashboard,
  LockKeyhole,
  LogOut,
  Menu,
  Plus,
  RefreshCw,
  Search,
  SlidersHorizontal,
  UsersRound,
  X,
} from 'lucide-react';

const frequencies = [
  'Once a Month',
  'Twice a Month',
  'Every Other Month',
  'Once a Week',
  'Twice a Week',
  'Every other Week',
];

const CENTRAL_TIME_ZONE = 'America/Chicago';

function centralDateKey(date = new Date()) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: CENTRAL_TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map(({ type, value }) => [type, value]));
  return `${values.year}-${values.month}-${values.day}`;
}

function dateValue(value) {
  if (!value) return null;
  const parsed = new Date(`${value.slice(0, 10)}T00:00:00Z`);
  return Number.isNaN(parsed.valueOf()) ? null : parsed;
}

function formatDate(value, options = { month: 'short', day: 'numeric', year: 'numeric' }) {
  const parsed = dateValue(value);
  return parsed ? new Intl.DateTimeFormat('en-US', { timeZone: 'UTC', ...options }).format(parsed) : '—';
}

function daysUntil(value, today = centralDateKey()) {
  const date = dateValue(value);
  if (!date) return null;
  return Math.round((date - new Date(`${today}T00:00:00`)) / 86400000);
}

function greeting(now) {
  const hour = Number(new Intl.DateTimeFormat('en-US', { timeZone: CENTRAL_TIME_ZONE, hour: 'numeric', hourCycle: 'h23' }).format(now));
  return hour < 12 ? 'Good morning' : hour < 18 ? 'Good afternoon' : 'Good evening';
}

function Metric({ label, value, note, tone, icon: Icon }) {
  return (
    <section className={`metric metric-${tone}`}>
      <div className="metric-top">
        <span>{label}</span>
        <Icon size={17} strokeWidth={1.8} />
      </div>
      <div className="metric-value">{value}</div>
      <div className="metric-note">{note}</div>
    </section>
  );
}

function Status({ value }) {
  const normalized = String(value || '').toLowerCase();
  const tone = normalized.includes('up to date') || normalized.includes('responsive')
    ? 'positive'
    : normalized.includes('pending') || normalized.includes('unresponsive')
      ? 'caution'
      : 'neutral';
  return <span className={`status status-${tone}`}><i />{value || 'Not set'}</span>;
}

function FilterMenu({ label, value, options, icon: Icon, onChange }) {
  const [open, setOpen] = useState(false);
  const menuId = useId();
  const rootRef = useRef(null);
  const triggerRef = useRef(null);
  const selectedOption = options.find((option) => option.value === value) || options[0];

  useEffect(() => {
    if (!open) return undefined;
    const closeOnOutsideClick = (event) => {
      if (!rootRef.current?.contains(event.target)) setOpen(false);
    };
    const closeOnEscape = (event) => {
      if (event.key !== 'Escape') return;
      setOpen(false);
      triggerRef.current?.focus();
    };
    document.addEventListener('pointerdown', closeOnOutsideClick);
    document.addEventListener('keydown', closeOnEscape);
    return () => {
      document.removeEventListener('pointerdown', closeOnOutsideClick);
      document.removeEventListener('keydown', closeOnEscape);
    };
  }, [open]);

  function focusOption(position) {
    const menuItems = rootRef.current?.querySelectorAll('.filter-option');
    menuItems?.[position]?.focus();
  }

  function onTriggerKeyDown(event) {
    if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
    event.preventDefault();
    setOpen(true);
    window.requestAnimationFrame(() => {
      const selectedIndex = options.findIndex((option) => option.value === value);
      focusOption(selectedIndex < 0 ? 0 : selectedIndex);
    });
  }

  function onOptionKeyDown(event, index) {
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      focusOption((index + 1) % options.length);
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      focusOption((index - 1 + options.length) % options.length);
    } else if (event.key === 'Home') {
      event.preventDefault();
      focusOption(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      focusOption(options.length - 1);
    }
  }

  return (
    <div className="filter-control filter-menu" ref={rootRef}>
      <button
        className="filter-trigger"
        type="button"
        ref={triggerRef}
        aria-label={`${label}: ${selectedOption.label}`}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-controls={menuId}
        onClick={() => setOpen((current) => !current)}
        onKeyDown={onTriggerKeyDown}
      >
        <Icon size={15} />
        <span>{selectedOption.label}</span>
        <ChevronDown size={14} />
      </button>
      {open && <div className="filter-menu-popover" id={menuId} role="menu" aria-label={label}>
        {options.map((option, index) => <button
          className="filter-option"
          type="button"
          role="menuitemradio"
          aria-checked={option.value === value}
          key={option.value}
          onClick={() => {
            onChange(option.value);
            setOpen(false);
            triggerRef.current?.focus();
          }}
          onKeyDown={(event) => onOptionKeyDown(event, index)}
        >
          <span>{option.label}</span>
          {option.value === value && <Check size={14} />}
        </button>)}
      </div>}
    </div>
  );
}

function ClientForm({ onClose, onCreated }) {
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  const [form, setForm] = useState({
    name: '',
    email: '',
    phone: '',
    translator: 'No',
    onboardingDate: centralDateKey(),
    frequency: 'Once a Month',
    interviewDate: '',
    lastContactDate: '',
    outcome: '',
    nextContactDate: '',
    lastAttorneyDate: '',
    paymentStatus: 'Pending',
    notes: '',
  });

  const update = (event) => setForm((current) => ({ ...current, [event.target.name]: event.target.value }));

  async function submit(event) {
    event.preventDefault();
    setSaving(true);
    setError('');
    try {
      const response = await fetch('/api/clients', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(form),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'Unable to add this client.');
      onCreated(result);
    } catch (submitError) {
      setError(submitError.message);
      setSaving(false);
    }
  }

  return (
    <div className="overlay" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <section className="drawer intake-drawer" role="dialog" aria-modal="true" aria-labelledby="intake-title">
        <div className="drawer-header">
          <div>
            <span className="eyebrow">NEW RECORD</span>
            <h2 id="intake-title">Client intake</h2>
          </div>
          <button className="icon-button" onClick={onClose} aria-label="Close intake"><X size={18} /></button>
        </div>
        <form onSubmit={submit} className="intake-form">
          <div className="form-section-title"><span>01</span><h3>Client details</h3></div>
          <div className="form-grid">
            <label className="field field-wide">Client name <b>*</b><input autoFocus name="name" value={form.name} onChange={update} required maxLength={120} placeholder="Full name" /></label>
            <label className="field">Email<input name="email" type="email" value={form.email} onChange={update} maxLength={180} placeholder="name@example.com" /></label>
            <label className="field">Phone<input name="phone" value={form.phone} onChange={update} maxLength={40} placeholder="Phone number" /></label>
            <label className="field">Meeting frequency<select name="frequency" value={form.frequency} onChange={update}>{frequencies.map((frequency) => <option key={frequency}>{frequency}</option>)}</select><ChevronDown size={15} /></label>
            <label className="field">Translator needed<select name="translator" value={form.translator} onChange={update}><option>No</option><option>Yes</option></select><ChevronDown size={15} /></label>
          </div>

          <div className="form-section-title"><span>02</span><h3>Case follow-up</h3></div>
          <div className="form-grid">
            <label className="field">Date onboarded<input type="date" name="onboardingDate" value={form.onboardingDate} onChange={update} /></label>
            <label className="field">Date of interview<input type="date" name="interviewDate" value={form.interviewDate} onChange={update} /></label>
            <label className="field">Date contacted<input type="date" name="lastContactDate" value={form.lastContactDate} onChange={update} /></label>
            <label className="field">Next contact<input type="date" name="nextContactDate" value={form.nextContactDate} onChange={update} /></label>
            <label className="field">Last date with attorney<input type="date" name="lastAttorneyDate" value={form.lastAttorneyDate} onChange={update} /></label>
            <label className="field">Outcome<select name="outcome" value={form.outcome} onChange={update}><option value="">Choose outcome</option><option>Responsive</option><option>Unresponsive</option></select><ChevronDown size={15} /></label>
            <label className="field">Payment status<select name="paymentStatus" value={form.paymentStatus} onChange={update}><option>Pending</option><option>Up to Date</option></select><ChevronDown size={15} /></label>
            <label className="field field-wide">Notes<textarea name="notes" value={form.notes} onChange={update} rows={3} maxLength={2000} placeholder="Add a note for this client" /></label>
          </div>
          {error && <div className="form-error" role="alert">{error}</div>}
          <div className="form-footer"><span>Saved to MasterData</span><button className="button button-primary" type="submit" disabled={saving}><Plus size={16} />{saving ? 'Saving…' : 'Add client'}</button></div>
        </form>
      </section>
    </div>
  );
}

function ClientDetails({ client, onClose }) {
  const fields = [
    ['Client ID', client.id], ['Email', client.email], ['Phone', client.phone],
    ['Meeting frequency', client.frequency], ['Translator needed', client.translator],
    ['Date onboarded', formatDate(client.onboardingDate)], ['Date of interview', formatDate(client.interviewDate)],
    ['Date contacted', formatDate(client.lastContactDate)], ['Outcome', client.outcome],
    ['Next date of contact', formatDate(client.nextContactDate)], ['Last date with attorney', formatDate(client.lastAttorneyDate)],
    ['Payment status', client.paymentStatus],
  ];
  return (
    <div className="overlay" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <section className="drawer detail-drawer" role="dialog" aria-modal="true" aria-labelledby="client-name">
        <div className="drawer-header">
          <div><span className="eyebrow">CLIENT RECORD · {client.id || 'NO ID'}</span><h2 id="client-name">{client.name}</h2></div>
          <button className="icon-button" onClick={onClose} aria-label="Close client details"><X size={18} /></button>
        </div>
        <div className="detail-status"><Status value={client.paymentStatus} /><span className="detail-frequency">{client.frequency || 'Frequency not set'}</span></div>
        <dl className="detail-grid">{fields.map(([label, value]) => <div key={label}><dt>{label}</dt><dd>{value || '—'}</dd></div>)}</dl>
        <section className="detail-notes"><h3>Notes</h3><p>{client.notes || 'No notes recorded.'}</p></section>
      </section>
    </div>
  );
}

function Dashboard({ clients, onShowClients, onSelectClient, now }) {
  const today = centralDateKey(now);
  const dueClients = useMemo(() => clients
    .map((client) => ({ ...client, dueIn: daysUntil(client.nextContactDate, today) }))
    .filter((client) => client.dueIn !== null)
    .sort((left, right) => left.dueIn - right.dueIn), [clients, today]);
  const overdue = dueClients.filter((client) => client.dueIn < 0).length;
  const dueSoon = dueClients.filter((client) => client.dueIn >= 0 && client.dueIn <= 7).length;
  const pendingPayments = clients.filter((client) => client.paymentStatus.toLowerCase() === 'pending').length;
  const nextFollowUp = dueClients.find((client) => client.dueIn >= 0) || null;
  const frequencyCounts = frequencies.map((frequency) => ({ frequency, count: clients.filter((client) => client.frequency.toLowerCase() === frequency.toLowerCase()).length }));
  const maxFrequency = Math.max(...frequencyCounts.map((item) => item.count), 1);
  const nextClients = dueClients.slice(0, 6);

  return (
    <>
      <div className="welcome-row"><div><span className="eyebrow">{new Intl.DateTimeFormat('en-US', { timeZone: CENTRAL_TIME_ZONE, weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }).format(now).toUpperCase()}</span><h1>{greeting(now)}.</h1><p className="page-subtitle">Here’s the client overview.</p></div><button className="button button-secondary" onClick={onShowClients}>View all clients <ArrowRight size={16} /></button></div>
      <div className="metric-grid">
        <Metric label="Total clients" value={clients.length} note="In MasterData" tone="green" icon={UsersRound} />
        <Metric label="Follow-ups due" value={overdue + dueSoon} note={`${overdue} overdue · ${dueSoon} in the next 7 days`} tone="amber" icon={CalendarClock} />
        <Metric label="Payments pending" value={pendingPayments} note={`${clients.length ? Math.round((pendingPayments / clients.length) * 100) : 0}% of client roster`} tone="rose" icon={CreditCard} />
        <Metric label="Next follow-up" value={nextFollowUp ? formatDate(nextFollowUp.nextContactDate, { month: 'short', day: 'numeric' }) : '—'} note={nextFollowUp?.name || 'No upcoming follow-up'} tone="blue" icon={CalendarClock} />
      </div>
      <div className="dashboard-grid">
        <section className="surface followup-surface">
          <div className="section-heading"><div><span className="eyebrow">CLIENT CARE</span><h2>Next follow-ups</h2></div><button className="text-button" onClick={onShowClients}>All clients <ArrowRight size={15} /></button></div>
          {nextClients.length ? <div className="followup-list">{nextClients.map((client) => <button className="followup-row" key={client.id || client.email} onClick={() => onSelectClient(client)}><span className={`due-marker ${client.dueIn < 0 ? 'due-overdue' : ''}`}><CalendarClock size={17} /></span><span className="followup-person"><strong>{client.name}</strong><small>{client.email || client.frequency || 'No contact details'}</small></span><span className={`due-date ${client.dueIn < 0 ? 'text-overdue' : ''}`}>{client.dueIn < 0 ? `${Math.abs(client.dueIn)}d overdue` : client.dueIn === 0 ? 'Today' : `In ${client.dueIn}d`}<small>{formatDate(client.nextContactDate, { month: 'short', day: 'numeric' })}</small></span><ArrowRight className="row-arrow" size={15} /></button>)}</div> : <div className="empty-state"><CalendarClock size={22} /><p>No follow-up dates recorded yet.</p></div>}
        </section>
        <section className="surface frequency-surface">
          <div className="section-heading"><div><span className="eyebrow">ROSTER MIX</span><h2>Meeting frequency</h2></div><span className="small-count">{clients.length} total</span></div>
          <div className="frequency-list">{frequencyCounts.map(({ frequency, count }, index) => <div className="frequency-row" key={frequency}><div className="frequency-label"><span>{frequency}</span><b>{count}</b></div><div className="bar-track"><i style={{ width: `${Math.max(count ? 4 : 0, (count / maxFrequency) * 100)}%`, '--bar-index': index }} /></div></div>)}</div>
          <div className="frequency-footer"><span><i />Active roster</span><span>{pendingPayments} payment pending</span></div>
        </section>
      </div>
    </>
  );
}

function ClientList({ clients, search, onSearch, onSelectClient }) {
  const [frequency, setFrequency] = useState('All frequencies');
  const [payment, setPayment] = useState('All payment statuses');
  const [sortOrder, setSortOrder] = useState('name');
  const visibleClients = clients.filter((client) => {
    const query = search.trim().toLowerCase();
    const matchesQuery = !query || [client.id, client.name, client.email, client.phone].some((value) => value.toLowerCase().includes(query));
    const matchesFrequency = frequency === 'All frequencies' || client.frequency === frequency;
    const matchesPayment = payment === 'All payment statuses' || client.paymentStatus === payment;
    return matchesQuery && matchesFrequency && matchesPayment;
  }).sort((left, right) => {
    if (sortOrder === 'clientId') {
      const leftNumber = Number(left.id.match(/\d+/)?.[0] || 0);
      const rightNumber = Number(right.id.match(/\d+/)?.[0] || 0);
      return leftNumber - rightNumber || left.name.localeCompare(right.name);
    }
    return left.name.localeCompare(right.name);
  });

  return (
    <>
      <div className="welcome-row roster-heading"><div><span className="eyebrow">MASTERDATA · {clients.length} RECORDS</span><h1>Client roster</h1><p className="page-subtitle">{visibleClients.length} clients shown</p></div></div>
      <section className="surface roster-surface">
        <div className="roster-toolbar">
          <label className="search-box"><Search size={17} /><input value={search} onChange={(event) => onSearch(event.target.value)} placeholder="Search name, email or ID" aria-label="Search clients" />{search && <button className="clear-search" onClick={() => onSearch('')} aria-label="Clear search"><X size={15} /></button>}</label>
          <FilterMenu label="Meeting frequency" value={frequency} onChange={setFrequency} icon={SlidersHorizontal} options={[{ value: 'All frequencies', label: 'All frequencies' }, ...frequencies.map((item) => ({ value: item, label: item }))]} />
          <FilterMenu label="Payment status" value={payment} onChange={setPayment} icon={CreditCard} options={[{ value: 'All payment statuses', label: 'All payment statuses' }, { value: 'Pending', label: 'Pending' }, { value: 'Up to Date', label: 'Up to Date' }]} />
          <FilterMenu label="Sort clients" value={sortOrder} onChange={setSortOrder} icon={ArrowDownAZ} options={[{ value: 'name', label: 'Name A-Z' }, { value: 'clientId', label: 'CL number' }]} />
        </div>
        <div className="table-scroll"><table className="client-table"><thead><tr><th>Client</th><th>Meeting frequency</th><th>Next contact</th><th>Payment</th><th aria-label="Open client" /></tr></thead><tbody>
          {visibleClients.map((client) => <tr key={client.id || client.email} onClick={() => onSelectClient(client)} tabIndex={0} onKeyDown={(event) => event.key === 'Enter' && onSelectClient(client)}><td><div className="client-name-cell"><span className="client-initials">{client.name.split(/\s+/).slice(0, 2).map((part) => part[0]).join('').toUpperCase()}</span><span><strong>{client.name}</strong><small>{client.id} <i /> {client.email || 'No email'}</small></span></div></td><td>{client.frequency || <span className="muted">Not set</span>}</td><td><span className={daysUntil(client.nextContactDate) < 0 ? 'text-overdue' : ''}>{formatDate(client.nextContactDate)}</span></td><td><Status value={client.paymentStatus} /></td><td><ArrowRight size={15} className="row-arrow" /></td></tr>)}
          {!visibleClients.length && <tr><td colSpan="5"><div className="table-empty"><Search size={22} /><strong>No matching clients</strong><span>Try another name or filter.</span></div></td></tr>}
        </tbody></table></div>
        <div className="table-foot"><span>Showing {visibleClients.length} of {clients.length}</span><span><i className="sync-dot" /> Synced from MasterData</span></div>
      </section>
    </>
  );
}

function LoginScreen({ onAuthenticated }) {
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');
  const [submitting, setSubmitting] = useState(false);

  async function submit(event) {
    event.preventDefault();
    setSubmitting(true);
    setError('');
    try {
      const response = await fetch('/api/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ password }),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'Sign-in failed.');
      onAuthenticated();
    } catch (loginError) {
      setError(loginError.message);
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <main className="login-screen">
      <section className="login-panel">
        <div className="login-mark"><LockKeyhole size={19} /></div>
        <span className="eyebrow">CLIENT SERVICES</span>
        <h1>Casework</h1>
        <p>Sign in to your client workspace.</p>
        <form onSubmit={submit}>
          <label className="field">Workspace password<span className="password-input"><input autoFocus type={showPassword ? 'text' : 'password'} value={password} onChange={(event) => setPassword(event.target.value)} required /><button className="password-toggle" type="button" onClick={() => setShowPassword((visible) => !visible)} aria-label={showPassword ? 'Hide password' : 'Show password'} aria-pressed={showPassword} title={showPassword ? 'Hide password' : 'Show password'}>{showPassword ? <EyeOff size={16} /> : <Eye size={16} />}</button></span></label>
          {error && <div className="form-error" role="alert">{error}</div>}
          <button className="button button-primary" type="submit" disabled={submitting}>{submitting ? 'Signing in…' : 'Sign in'} <ArrowRight size={15} /></button>
        </form>
      </section>
    </main>
  );
}

function CentralClock() {
  const [now, setNow] = useState(() => new Date());

  useEffect(() => {
    const timer = window.setInterval(() => setNow(new Date()), 1000);
    return () => window.clearInterval(timer);
  }, []);

  const date = new Intl.DateTimeFormat('en-US', { timeZone: CENTRAL_TIME_ZONE, weekday: 'short', month: 'short', day: 'numeric' }).format(now);
  const time = new Intl.DateTimeFormat('en-US', { timeZone: CENTRAL_TIME_ZONE, hour: 'numeric', minute: '2-digit', second: '2-digit', timeZoneName: 'short' }).format(now);

  return <span className="central-clock"><Clock3 size={15} /><span><b>{date}</b><time>{time}</time></span></span>;
}

export default function App() {
  const [page, setPage] = useState('dashboard');
  const [clients, setClients] = useState([]);
  const [search, setSearch] = useState('');
  const [selectedClient, setSelectedClient] = useState(null);
  const [intakeOpen, setIntakeOpen] = useState(false);
  const [loadError, setLoadError] = useState('');
  const [lastUpdated, setLastUpdated] = useState(null);
  const [dataSource, setDataSource] = useState('Excel');
  const [refreshing, setRefreshing] = useState(false);
  const [toast, setToast] = useState('');
  const [authStatus, setAuthStatus] = useState('checking');
  const [loginRequired, setLoginRequired] = useState(false);
  const [, setCentralDay] = useState(() => centralDateKey());

  useEffect(() => {
    let active = true;
    fetch('/api/session')
      .then((response) => response.json())
      .then((result) => {
        if (!active) return;
        setLoginRequired(result.loginRequired);
        setAuthStatus(result.authenticated ? 'authenticated' : 'anonymous');
      })
      .catch(() => active && setAuthStatus('anonymous'));
    return () => { active = false; };
  }, []);

  useEffect(() => {
    const timer = window.setInterval(() => {
      const today = centralDateKey();
      setCentralDay((current) => current === today ? current : today);
    }, 30000);
    return () => window.clearInterval(timer);
  }, []);

  useEffect(() => {
    if (authStatus !== 'authenticated') return undefined;
    let active = true;
    const load = async () => {
      try {
        const response = await fetch('/api/clients');
        const result = await response.json();
        if (!response.ok) throw new Error(result.error || 'Could not load the client data.');
        if (active) {
          setClients(result.clients);
          setLastUpdated(result.modifiedAt);
          setDataSource(result.source || 'Excel');
          setLoadError('');
        }
      } catch (error) {
        if (active) setLoadError(error.message);
      }
    };
    load();
    const timer = window.setInterval(load, 20000);
    window.addEventListener('focus', load);
    return () => {
      active = false;
      window.clearInterval(timer);
      window.removeEventListener('focus', load);
    };
  }, [authStatus]);

  useEffect(() => {
    if (!toast) return undefined;
    const timer = window.setTimeout(() => setToast(''), 3600);
    return () => window.clearTimeout(timer);
  }, [toast]);

  async function refreshWorkbook() {
    setRefreshing(true);
    try {
      const response = await fetch('/api/clients');
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'Could not refresh client data.');
      setClients(result.clients);
      setLastUpdated(result.modifiedAt);
      setDataSource(result.source || 'Excel');
      setLoadError('');
      setToast('Client data refreshed');
    } catch (error) {
      setLoadError(error.message);
    } finally {
      setRefreshing(false);
    }
  }

  function createdClient(result) {
    setClients((current) => [...current, result.client].sort((left, right) => left.name.localeCompare(right.name)));
    setLastUpdated(result.modifiedAt);
    setDataSource(result.source || dataSource);
    setIntakeOpen(false);
    setPage('clients');
    setSearch(result.client.name);
    setToast(`${result.client.name} added to MasterData`);
  }

  async function logout() {
    await fetch('/api/logout', { method: 'POST' });
    setAuthStatus('anonymous');
  }

  const updatedLabel = lastUpdated
    ? `Synced ${new Intl.DateTimeFormat('en-US', { timeZone: CENTRAL_TIME_ZONE, hour: 'numeric', minute: '2-digit' }).format(new Date(lastUpdated))} CT`
    : 'Waiting for data';
  if (authStatus === 'checking') return <main className="login-screen"><div className="login-mark"><LockKeyhole size={19} /></div></main>;
  if (authStatus !== 'authenticated') return <LoginScreen onAuthenticated={() => setAuthStatus('authenticated')} />;

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <a className="brand" href="#dashboard" onClick={(event) => { event.preventDefault(); setPage('dashboard'); }}><span className="brand-mark"><span /><span /><span /></span><span className="brand-name">casework<span>CLIENT SERVICES</span></span></a>
        <div className="sidebar-section"><span className="nav-caption">WORKSPACE</span><button className={`nav-item ${page === 'dashboard' ? 'active' : ''}`} onClick={() => setPage('dashboard')}><LayoutDashboard size={18} /><span>Overview</span></button><button className={`nav-item ${page === 'clients' ? 'active' : ''}`} onClick={() => setPage('clients')}><UsersRound size={18} /><span>Clients</span><span className="nav-count">{clients.length}</span></button></div>
        <div className="sidebar-bottom"><div className="workbook-chip"><FileSpreadsheet size={17} /><span><b>MasterData</b><small>{dataSource === 'Google Sheets' ? 'Google Sheets' : 'Excel workbook'}</small></span><span className="live-mark" title="Live data connection" /></div><div className="sidebar-user"><span className="user-avatar">CS</span><span><b>Case services</b><small>{loginRequired ? 'Private workspace' : 'Local workspace'}</small></span><Menu size={17} /></div></div>
      </aside>
      <main className="main-area">
        <header className="topbar"><div className="breadcrumb"><span>Workspace</span><span>/</span><b>{page === 'dashboard' ? 'Overview' : 'Clients'}</b></div><div className="topbar-actions"><CentralClock /><span className="sync-state"><i className="sync-dot" />{updatedLabel}</span><button className={`icon-button refresh-button ${refreshing ? 'is-refreshing' : ''}`} onClick={refreshWorkbook} aria-label="Refresh client data" title="Refresh data"><RefreshCw size={17} /></button>{loginRequired && <button className="icon-button" onClick={logout} aria-label="Sign out" title="Sign out"><LogOut size={16} /></button>}<button className="button button-primary top-add" onClick={() => setIntakeOpen(true)}><Plus size={16} />Add client</button></div></header>
        <div className="page-content">
          {loadError && <div className="load-error" role="alert"><span>{loadError}</span><button onClick={refreshWorkbook}>Retry <RefreshCw size={14} /></button></div>}
          {page === 'dashboard'
            ? <Dashboard clients={clients} now={new Date()} onShowClients={() => setPage('clients')} onSelectClient={setSelectedClient} />
            : <ClientList clients={clients} search={search} onSearch={setSearch} onSelectClient={setSelectedClient} />}
          <footer className="page-footer"><span><i className="sync-dot" /> {dataSource} connection active</span><span>{dataSource === 'Google Sheets' ? 'Sheet checked every 20 seconds' : 'Workbook checked every 20 seconds'}</span></footer>
        </div>
      </main>
      {intakeOpen && <ClientForm onClose={() => setIntakeOpen(false)} onCreated={createdClient} />}
      {selectedClient && <ClientDetails client={selectedClient} onClose={() => setSelectedClient(null)} />}
      {toast && <div className="toast"><span><Check size={15} /></span>{toast}</div>}
    </div>
  );
}
