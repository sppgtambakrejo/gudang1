import { FormEvent, ReactNode, useEffect, useState } from 'react'
import { LockKeyhole, LogOut, Warehouse } from 'lucide-react'
import { isSupabaseConfigured, supabase } from './supabase'

export default function OwnerAuth({ children }: { children: ReactNode }) {
  const [loading, setLoading] = useState(true)
  const [authenticated, setAuthenticated] = useState(false)
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [submitting, setSubmitting] = useState(false)

  useEffect(() => {
    if (!isSupabaseConfigured) {
      setLoading(false)
      return
    }
    supabase.auth.getSession().then(({ data }) => {
      setAuthenticated(Boolean(data.session))
      setLoading(false)
    })
    const { data } = supabase.auth.onAuthStateChange((_event, session) => {
      setAuthenticated(Boolean(session))
    })
    return () => data.subscription.unsubscribe()
  }, [])

  async function login(event: FormEvent) {
    event.preventDefault()
    setSubmitting(true)
    setError('')
    const { error: loginError } = await supabase.auth.signInWithPassword({ email, password })
    if (loginError) setError('Email atau password pemilik tidak benar.')
    setSubmitting(false)
  }

  if (loading) return <div style={styles.center}>Memuat aplikasi…</div>

  if (!isSupabaseConfigured) {
    return (
      <div style={styles.center}>
        <div style={styles.card}>
          <Warehouse size={38} />
          <h2>Supabase belum dikonfigurasi</h2>
          <p>Isi VITE_SUPABASE_URL dan VITE_SUPABASE_PUBLISHABLE_KEY pada file .env atau GitHub Actions Secrets.</p>
        </div>
      </div>
    )
  }

  if (!authenticated) {
    return (
      <div style={styles.center}>
        <form style={styles.card} onSubmit={login}>
          <LockKeyhole size={38} />
          <h2>Login Pemilik SPPG</h2>
          <p>Masuk menggunakan akun email yang dibuat di Supabase Authentication.</p>
          <input style={styles.input} type="email" placeholder="Email pemilik" value={email} onChange={(e) => setEmail(e.target.value)} required />
          <input style={styles.input} type="password" placeholder="Password" value={password} onChange={(e) => setPassword(e.target.value)} required />
          {error && <div style={styles.error}>{error}</div>}
          <button style={styles.button} disabled={submitting}>{submitting ? 'Memproses…' : 'Masuk'}</button>
        </form>
      </div>
    )
  }

  return (
    <>
      <button style={styles.logout} onClick={() => supabase.auth.signOut()} title="Keluar dari akun pemilik">
        <LogOut size={16} /> Keluar Pemilik
      </button>
      {children}
    </>
  )
}

const styles: Record<string, React.CSSProperties> = {
  center: { minHeight: '100vh', display: 'grid', placeItems: 'center', background: '#f2f4f0', padding: 20, fontFamily: 'Arial, sans-serif' },
  card: { width: '100%', maxWidth: 390, background: '#fff', border: '1px solid #d9ded7', borderRadius: 18, padding: 28, display: 'grid', gap: 14, boxShadow: '0 18px 50px rgba(30,50,35,.10)' },
  input: { padding: '13px 14px', borderRadius: 10, border: '1px solid #cfd6cd', fontSize: 15 },
  button: { padding: '13px 14px', border: 0, borderRadius: 10, background: '#2b6e52', color: '#fff', fontWeight: 700, cursor: 'pointer' },
  error: { background: '#fff0ef', color: '#9d2f26', borderRadius: 8, padding: 10, fontSize: 13 },
  logout: { position: 'fixed', right: 12, bottom: 12, zIndex: 1000, display: 'flex', alignItems: 'center', gap: 6, border: '1px solid #d7ddd5', background: '#fff', borderRadius: 999, padding: '8px 12px', cursor: 'pointer', boxShadow: '0 5px 20px rgba(0,0,0,.12)' },
}
