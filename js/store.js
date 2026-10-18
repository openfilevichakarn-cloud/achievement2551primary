import { SUPABASE_URL, SUPABASE_ANON_KEY } from './config.js';

const remote = !!(SUPABASE_URL && SUPABASE_ANON_KEY && window.supabase);
const sb = remote ? window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;
const LS = 'pp5_2551_';
const lsGet = (k, d) => { try { return JSON.parse(localStorage.getItem(LS + k)) ?? d; } catch { return d; } };
const lsSet = (k, v) => localStorage.setItem(LS + k, JSON.stringify(v));

const timers = {};
function debounce(key, fn, ms = 900) {
  clearTimeout(timers[key]);
  return new Promise((res, rej) => { timers[key] = setTimeout(() => fn().then(res, rej), ms); });
}

export const store = {
  remote,
  async user() {
    if (!remote) return { id: 'local', email: 'โหมดทดลอง' };
    const { data } = await sb.auth.getSession();
    return data.session?.user ?? null;
  },
  async signIn(email, password) {
    const { error } = await sb.auth.signInWithPassword({ email, password });
    if (error) throw error;
  },
  async signUp(email, password) {
    const { data, error } = await sb.auth.signUp({ email, password });
    if (error) throw error;
    return !!data.session;
  },
  async signOut() { if (remote) await sb.auth.signOut(); },

  async loadSchool() {
    if (!remote) return lsGet('school', {});
    const { data, error } = await sb.from('school_settings').select('data').maybeSingle();
    if (error) throw error;
    return data?.data ?? {};
  },
  saveSchool(v) {
    if (!remote) { lsSet('school', v); return Promise.resolve(); }
    return debounce('school', async () => {
      const { error } = await sb.from('school_settings').upsert({ data: v, updated_at: new Date().toISOString() });
      if (error) throw error;
    });
  },

  async listRooms() {
    if (!remote) return lsGet('rooms', []).map(r => ({ id: r.id, title: r.title, year: r.year }));
    const { data, error } = await sb.from('classrooms').select('id,title,year').order('year', { ascending: false }).order('title');
    if (error) throw error;
    return data;
  },
  async loadRoom(id) {
    if (!remote) return lsGet('rooms', []).find(r => r.id === id) ?? null;
    const { data, error } = await sb.from('classrooms').select('id,title,year,data').eq('id', id).maybeSingle();
    if (error) throw error;
    return data ? { ...data.data, id: data.id, title: data.title, year: data.year } : null;
  },
  async createRoom(room) {
    if (!remote) {
      room.id = crypto.randomUUID();
      const all = lsGet('rooms', []); all.push(room); lsSet('rooms', all);
      return room;
    }
    const { id, title, year, ...rest } = room;
    const { data, error } = await sb.from('classrooms').insert({ title, year, data: rest }).select('id').single();
    if (error) throw error;
    room.id = data.id;
    return room;
  },
  saveRoom(room) {
    if (!remote) {
      const all = lsGet('rooms', []);
      const i = all.findIndex(r => r.id === room.id);
      if (i >= 0) all[i] = room; else all.push(room);
      lsSet('rooms', all);
      return Promise.resolve();
    }
    return debounce('room' + room.id, async () => {
      const { id, title, year, ...rest } = room;
      const { error } = await sb.from('classrooms').update({ title, year, data: rest, updated_at: new Date().toISOString() }).eq('id', id);
      if (error) throw error;
    });
  },
  async deleteRoom(id) {
    if (!remote) { lsSet('rooms', lsGet('rooms', []).filter(r => r.id !== id)); return; }
    const { error } = await sb.from('classrooms').delete().eq('id', id);
    if (error) throw error;
  },
};
