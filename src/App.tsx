
import React, { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

type Tipo = 'Físico' | 'Mágico' | 'Buff'
type StatKey = 'str' | 'defFis' | 'defMag' | 'mag' | 'spd' | 'risk'
type Mods = Partial<Record<StatKey, number>>
type HpMode = 'Barrera' | 'Curación' | ''

interface Ability {
  name: string
  type: Tipo
  doesDamage: boolean
  bonusDamage: number
  target: 'Atacante' | 'Defensor'
  mods: Mods
  modRL: number
  modHP: number
  hpMode: HpMode
  allowedFor: string[]
  duration?: number | null
}
interface BaseStats {
  vida: number
  str: number
  defFis: number
  defMag: number
  mag: number
  spd: number
  risk: number
}
interface Turn {
  attacker: string
  defender: string
  abilityName: string
  type: Tipo
  baseDamage: number | null
  defenseApplied: number | null
  damagePlusAbility: number | null
  finalDamage: number
}
interface ActiveEffect {
  id: string
  owner: string
  abilityName: string
  mods: Mods
  remaining: number | null
  barrierRemaining?: number | null
  barrierMax?: number | null
  createdAt: number
}
interface HealEvent {
  character: string
  amount: number
  at: number
}
type Dataset = { characters: Record<string, BaseStats>, abilities: Ability[] }

const DEFAULT_DATA: Dataset = {
  characters: {
    Isildur: { vida: 190, str: 3, defFis: 1, defMag: 10, mag: 26, spd: 3, risk: 1 },
    Zyhna:   { vida: 220, str: 5, defFis: 8, defMag: 2,  mag: 12, spd: 4, risk: 2 },
    Henrich: { vida: 250, str: 8, defFis: 6, defMag: 5,  mag: 6,  spd: 2, risk: 3 },
  },
  abilities: [
    { name:'Golpe', type:'Físico', doesDamage:true,  bonusDamage:1,  target:'Atacante', mods:{}, modRL:0, modHP:0, hpMode:'', allowedFor:['*'], duration: null },
    { name:'Cañón mágico', type:'Mágico', doesDamage:true,  bonusDamage:10, target:'Atacante', mods:{}, modRL:0, modHP:0, hpMode:'', allowedFor:['Zyhna','Isildur'], duration: null },
    { name:'Bola de fuego', type:'Mágico', doesDamage:true,  bonusDamage:3,  target:'Atacante', mods:{}, modRL:0, modHP:0, hpMode:'', allowedFor:['Zyhna'], duration: null },
    { name:'Jihn', type:'Mágico', doesDamage:true,  bonusDamage:20, target:'Atacante', mods:{}, modRL:0, modHP:0, hpMode:'', allowedFor:['Isildur'], duration: null },
    { name:'Viento oscuro', type:'Mágico', doesDamage:true,  bonusDamage:10, target:'Defensor',  mods:{ mag:-5 }, modRL:0, modHP:0, hpMode:'', allowedFor:['*'], duration: 3 },
    { name:'Emoción eléctrica', type:'Buff',   doesDamage:false, bonusDamage:0,  target:'Atacante', mods:{ mag:5, spd:5 }, modRL:0, modHP:0, hpMode:'', allowedFor:['*'], duration: 2 },
    { name:'Armadura gélida', type:'Buff', doesDamage:false, bonusDamage:0, target:'Atacante', mods:{ defFis:10, defMag:10 }, modRL:0, modHP:40, hpMode:'Barrera', allowedFor:['*'], duration: 3 },
    { name:'Curación menor', type:'Buff', doesDamage:false, bonusDamage:0, target:'Atacante', mods:{}, modRL:0, modHP:20, hpMode:'Curación', allowedFor:['*'], duration: null },
    { name:'Valentía', type:'Buff', doesDamage:false, bonusDamage:0, target:'Atacante', mods:{ spd:2, risk:2 }, modRL:2, modHP:0, hpMode:'', allowedFor:['*'], duration: 2 },
  ]
}

function addMods(a: Mods, b: Mods): Mods {
  const out: Mods = { ...a }
  ;(['str','defFis','defMag','mag','spd','risk'] as StatKey[]).forEach(k => {
    if (b[k] !== undefined) out[k] = (out[k] ?? 0) + (b[k] ?? 0)
  })
  return out
}
const toBool = (x: any) => {
  if (typeof x === 'boolean') return x
  const s = String(x ?? '').trim().toLowerCase()
  return ['1','true','sí','si','y','yes','x'].includes(s)
}
const toNum = (x: any) => {
  if (typeof x === 'number') return x
  const n = parseFloat(String(x ?? '').replace(',','.'))
  return Number.isFinite(n) ? n : 0
}
const toInt = (x: any) => {
  if (x === null || x === undefined || x === '') return null
  const n = parseInt(String(x).trim(), 10)
  return Number.isFinite(n) ? n : null
}

function parseCharacters(ws: XLSX.WorkSheet): Record<string, BaseStats> {
  const rows: any[] = XLSX.utils.sheet_to_json(ws, { defval: '' })
  const out: Record<string, BaseStats> = {}
  for (const row of rows) {
    const name = String(row['Nombre'] ?? row['nombre'] ?? row['NOMBRE'] ?? '').trim()
    if (!name) continue
    const vida = toNum(row['Vida'] ?? row['vida'])
    const fuerza = toNum(row['Fuerza'] ?? row['fuerza'])
    const defFis = toNum(row['Def. Física'] ?? row['Def. Fisica'] ?? row['def fisica'] ?? row['Def Física'] ?? row['Defensa Física'] ?? row['Defensa Fisica'])
    const defMag = toNum(row['Def. Mágica'] ?? row['Def. Magica'] ?? row['def magica'] ?? row['Def Mágica'] ?? row['Defensa Mágica'] ?? row['Defensa Magica'])
    const magia = toNum(row['Magia'] ?? row['magia'])
    const vel = toNum(row['Velocidad'] ?? row['velocidad'])
    const risk = toNum(row['Nivel de riesgo'] ?? row['Riesgo'] ?? row['nivel de riesgo'] ?? 0)
    out[name] = { vida, str: fuerza, defFis, defMag, mag: magia, spd: vel, risk }
  }
  return out
}

function parseAbilities(ws: XLSX.WorkSheet): Ability[] {
  const rows: any[] = XLSX.utils.sheet_to_json(ws, { defval: '' })
  const out: Ability[] = []
  for (const row of rows) {
    const name = String(row['Habilidad'] ?? row['habilidad'] ?? '').trim()
    if (!name) continue
    const tipoS = String(row['Tipo'] ?? row['tipo'] ?? '').trim() as Tipo
    const tipo: Tipo = (tipoS === 'Físico' || tipoS === 'Mágico' || tipoS === 'Buff') ? tipoS : 'Físico'
    const doesDamage = toBool(row['Hace Daño'] ?? row['Hace Dano'] ?? row['Damage'] ?? row['hace daño'])
    const bonusDamage = toNum(row['Bonus Daño'] ?? row['Bonus Dano'] ?? row['bonus'] ?? 0)
    const tgtS = String(row['Objetivo'] ?? row['objetivo'] ?? '').trim()
    const target: 'Atacante' | 'Defensor' = (tgtS === 'Defensor') ? 'Defensor' : 'Atacante'

    const mods: Mods = {
      str: toNum(row['Mod Fuerza'] ?? row['mod fuerza'] ?? 0),
      defFis: toNum(row['Mod Def. Física'] ?? row['mod def. física'] ?? row['mod def fisica'] ?? 0),
      defMag: toNum(row['Mod Def. Mágica'] ?? row['mod def. mágica'] ?? row['mod def magica'] ?? 0),
      mag: toNum(row['Mod Magia'] ?? row['mod magia'] ?? 0),
      spd: toNum(row['Mod Velocidad'] ?? row['mod velocidad'] ?? 0),
      risk: toNum(row['Mod RL'] ?? row['mod rl'] ?? 0),
    }
    const modHP = toNum(row['Mod HP'] ?? row['mod hp'] ?? 0)
    const hpModeS = String(row['HP Tipo'] ?? row['Tipo HP'] ?? row['HP Modo'] ?? '').trim()
    const hpMode: HpMode = (hpModeS === 'Barrera' || hpModeS === 'Curación') ? hpModeS : ''

    const allowedCol = String(row['AllowedFor'] ?? row['Permitidos'] ?? '').trim()
    const allowedFor = allowedCol ? allowedCol.split(/[;,]/).map(s => s.trim()).filter(Boolean) : ['*']
    const duration = toInt(row['Duración'] ?? row['duración'] ?? row['Duration'] ?? row['duracion'])

    out.push({ name, type: tipo, doesDamage, bonusDamage, target, mods, modRL: mods.risk ?? 0, modHP, hpMode, allowedFor, duration })
  }
  return out
}

function applyListas(ws: XLSX.WorkSheet, abilities: Ability[]) {
  const rows: any[] = XLSX.utils.sheet_to_json(ws, { defval: '' , header: 1 })
  if (!rows.length) return
  let pi = -1, hi = -1
  const header = rows[0].map((x:any)=>String(x))
  const norm = (h: string) => h.normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[.]/g,'').trim().toLowerCase()
  for (let i=0;i<header.length;i++){
    const h = norm(header[i] ?? '')
    if (h === 'personaje') pi = i
    if (h.startsWith('habilidad')) hi = i
  }
  let currentP = ''
  const mapping: Record<string, Set<string>> = {}
  for (let r=1; r<rows.length; r++) {
    const row = rows[r] as any[]
    if (pi >= 0) {
      const pcell = String(row[pi] ?? '').trim()
      if (pcell) currentP = pcell
    }
    const hcell = hi >= 0 ? String(row[hi] ?? '').trim() : ''
    if (currentP && hcell) {
      if (!mapping[hcell]) mapping[hcell] = new Set()
      mapping[hcell].add(currentP)
    }
  }
  abilities.forEach(a => {
    const set = mapping[a.name]
    if (set && set.size) a.allowedFor = Array.from(set)
  })
}

export default function App() {
  const [data, setData] = useState<Dataset>(DEFAULT_DATA)
  const [turns, setTurns] = useState<Turn[]>([])
  const [effects, setEffects] = useState<ActiveEffect[]>([])
  const [heals, setHeals] = useState<HealEvent[]>([])
  const [message, setMessage] = useState<string>('')
  const [error, setError] = useState<string>('')

  const [hpAdjust, setHpAdjust] = useState<Record<string, number>>({})
  const [selected, setSelected] = useState<Record<string, boolean>>({})

  const names = useMemo(()=>Object.keys(data.characters), [data.characters])
  const selectedNames = useMemo(()=> names.filter(n => selected[n]), [names, selected])

  const [attacker, setAttacker] = useState<string>('')
  const [defender, setDefender] = useState<string>('')
  const [abilityName, setAbilityName] = useState<string>('')

  // Custom action state (Daño extra)
  const [cName, setCName] = useState<string>('Custom')
  const [cDoesDamage, setCDoesDamage] = useState<boolean>(true)
  const [cType, setCType] = useState<Tipo>('Físico')
  const [cExtraDamage, setCExtraDamage] = useState<number>(0)
  const [cTarget, setCTarget] = useState<'Atacante'|'Defensor'>('Atacante')
  const [cMods, setCMods] = useState<Mods>({ str:0, defFis:0, defMag:0, mag:0, spd:0, risk:0 })
  const [cHPMode, setCHPMode] = useState<HpMode>('')
  const [cHP, setCHP] = useState<number>(0)
  const [cDur, setCDur] = useState<number | null>(null)

  useEffect(() => {
    const nn = Object.keys(data.characters)
    const sel: Record<string, boolean> = {}
    const hadj: Record<string, number> = {}
    nn.forEach(n => { sel[n] = true; hadj[n] = 0 })
    setSelected(sel)
    setHpAdjust(hadj)
    setAttacker(nn[0] || '')
    setDefender(nn[1] || nn[0] || '')
    setEffects([])
    setHeals([])
  }, [data])

  useEffect(() => {
    if (!selectedNames.includes(attacker)) setAttacker(selectedNames[0] || '')
    if (!selectedNames.includes(defender)) setDefender(selectedNames[0] || '')
  }, [selectedNames])

  const ability = useMemo(
    () => data.abilities.find(a => a.name === abilityName) || null,
    [abilityName, data.abilities]
  )
  const allowedAbilities = useMemo(() => {
    if (!attacker) return data.abilities
    return data.abilities.filter(a => a.allowedFor.includes('*') || a.allowedFor.includes(attacker))
  }, [attacker, data.abilities])

  function addMods(a: Mods, b: Mods): Mods {
    const out: Mods = { ...a }
    ;(['str','defFis','defMag','mag','spd','risk'] as StatKey[]).forEach(k => {
      if (b[k] !== undefined) out[k] = (out[k] ?? 0) + (b[k] ?? 0)
    })
    return out
  }

  function currentModsFor(character: string): Mods {
    let m: Mods = {}
    for (const ef of effects) {
      if (ef.owner !== character) continue
      if (!(ef.remaining === null || ef.remaining > 0)) continue
      const barrierOk = (ef.barrierRemaining === undefined) || ((ef.barrierRemaining ?? 0) > 0)
      if (barrierOk) m = addMods(m, ef.mods)
    }
    return m
  }
  function effectiveStat(character: string, key: StatKey) {
    const base = data.characters[character][key]
    const delta = currentModsFor(character)[key] ?? 0
    const val = base + delta
    return key === 'risk' ? val : Math.max(0, val)  // RL puede ser negativo
  }
  function totalDamageReceived(character: string, upToIndexInclusive: number) {
    let total = 0
    for (let i=0;i<=upToIndexInclusive;i++){
      const t = turns[i]
      if (!t) continue
      if (t.defender === character) total += t.finalDamage
    }
    return total
  }
  function totalHealed(character: string, upToIndexInclusive: number) {
    let total = 0
    for (const h of heals) if (h.character === character && h.at <= upToIndexInclusive) total += h.amount
    return total
  }
  function currentHP(character: string, upToIndexInclusive: number) {
    const adj = hpAdjust[character] ?? 0
    return Math.max(0, data.characters[character].vida - totalDamageReceived(character, upToIndexInclusive) + totalHealed(character, upToIndexInclusive) + adj)
  }
  function totalBarrier(character: string) {
    let s = 0
    for (const ef of effects) {
      if (ef.owner === character && (ef.remaining === null || ef.remaining > 0) && (ef.barrierRemaining ?? 0) > 0) {
        s += (ef.barrierRemaining ?? 0)
      }
    }
    return s
  }
  const turnOrder = useMemo(() => {
    const list = selectedNames.map(n => ({ name: n, spd: effectiveStat(n, 'spd') }))
    list.sort((a,b) => (b.spd - a.spd) || a.name.localeCompare(b.name))
    return list
  }, [selectedNames, effects, data])

  function pushEffect(ability: Ability, owner: string) {
    const hasMods = Object.values(ability.mods).some(v => (v ?? 0) !== 0)
    const isBarrier = (ability.hpMode === 'Barrera') && (ability.modHP > 0)
    if (!hasMods && !isBarrier) return
    const ef: ActiveEffect = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2,8)}`,
      owner,
      abilityName: ability.name,
      mods: ability.mods,
      remaining: ability.duration ?? null,
      barrierRemaining: isBarrier ? ability.modHP : undefined,
      barrierMax: isBarrier ? ability.modHP : undefined,
      createdAt: turns.length
    }
    setEffects(prev => prev.concat([ef]))
  }

  function consumeBarriers(defender: string, incoming: number) {
    let remaining = incoming
    let absorbed = 0
    const updates = new Map<string, number>()
    for (const ef of effects) {
      if (ef.owner !== defender) continue
      if (!(ef.remaining === null || ef.remaining > 0)) continue
      const br = ef.barrierRemaining ?? 0
      if (br <= 0) continue
      const d = Math.min(br, remaining)
      if (d > 0) {
        updates.set(ef.id, br - d)
        remaining -= d
        absorbed += d
        if (remaining <= 0) break
      }
    }
    if (updates.size) {
      setEffects(prev => prev.map(e => updates.has(e.id) ? { ...e, barrierRemaining: updates.get(e.id)! } : e))
    }
    return { absorbed, remaining }
  }

  function tickAttackerBuffs(attacker: string) {
    setEffects(prev => prev.map(e => {
      if (e.owner === attacker && e.remaining !== null && e.remaining > 0) {
        return { ...e, remaining: e.remaining - 1 }
      }
      return e
    }).filter(e => e.remaining === null || e.remaining > 0))
  }

  function applyAbility(ability: Ability, attacker: string, defender: string, customBase: number | null = null) {
    let baseDamage: number | null = null
    let defenseApplied: number | null = null
    let damagePlusAbility: number | null = null
    let finalDamage = 0

    if (ability.doesDamage) {
      if (customBase !== null) {
        baseDamage = Math.max(0, customBase)
      } else {
        if (ability.type === 'Físico') {
          baseDamage = effectiveStat(attacker, 'str')
          defenseApplied = effectiveStat(defender, 'defFis')
        } else if (ability.type === 'Mágico') {
          baseDamage = effectiveStat(attacker, 'mag')
          defenseApplied = effectiveStat(defender, 'defMag')
        } else {
          // Para tipos no Físico/Mágico (p.ej., Buff) cuando se marcó "Hace daño":
          baseDamage = 0
          defenseApplied = 0
        }
      }
      if (defenseApplied === null) {
        defenseApplied = ability.type === 'Físico' ? effectiveStat(defender, 'defFis')
                          : ability.type === 'Mágico' ? effectiveStat(defender, 'defMag')
                          : 0
      }
      const raw = Math.max(0, ((baseDamage ?? 0) + (ability.bonusDamage || 0)) - (defenseApplied ?? 0))
      const { absorbed, remaining } = consumeBarriers(defender, raw)
      damagePlusAbility = (baseDamage ?? 0) + (ability.bonusDamage || 0)
      finalDamage = Math.max(0, remaining)

      tickAttackerBuffs(attacker)
      const targetName = ability.target === 'Atacante' ? attacker : defender
      pushEffect(ability, targetName)
      if (ability.hpMode === 'Curación' && ability.modHP > 0) {
        setHeals(prev => prev.concat([{ character: targetName, amount: ability.modHP, at: turns.length }]))
      }
    } else {
      tickAttackerBuffs(attacker)
      const targetName = ability.target === 'Atacante' ? attacker : defender
      pushEffect(ability, targetName)
      if (ability.hpMode === 'Curación' && ability.modHP > 0) {
        setHeals(prev => prev.concat([{ character: targetName, amount: ability.modHP, at: turns.length }]))
      }
    }

    const t: Turn = { attacker, defender, abilityName: ability.name, type: ability.type, baseDamage, defenseApplied, damagePlusAbility, finalDamage }
    setTurns(prev => [...prev, t])
  }

  function computeAndAddTurn() {
    if (!attacker || !defender || !ability) return
    applyAbility(ability, attacker, defender, null)
    setAbilityName('')
  }

  function computeCustomTurn() {
    if (!attacker || !defender) return
    const customAbility: Ability = {
      name: cName || 'Custom',
      type: cType,
      doesDamage: cDoesDamage,
      bonusDamage: cExtraDamage || 0, // "Daño extra"
      target: cTarget,
      mods: { ...cMods },
      modRL: cMods.risk ?? 0,
      modHP: cHP || 0,
      hpMode: cHPMode,
      allowedFor: ['*'],
      duration: cDur,
    }
    // No pasamos base custom: se toma de STR/MAG según tipo
    applyAbility(customAbility, attacker, defender, null)
    // Reset suave
    setCName('Custom'); setCExtraDamage(0); setCHP(0)
  }

  function resetAll(){ setTurns([]); setEffects([]); setHeals([]) }
  function undoLast(){ setTurns(prev => prev.slice(0,-1)) }
  function clearEffectsFor(name: string) { setEffects(prev => prev.filter(e => e.owner !== name)) }
  function clearAllEffects() { setEffects([]) }

  async function onExcelPicked(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    setError(''); setMessage('')
    try {
      const buf = await file.arrayBuffer()
      const wb = XLSX.read(buf, { type: 'array' })
      const wsP = wb.Sheets['Personajes']
      const wsH = wb.Sheets['Habilidades']
      if (!wsP || !wsH) throw new Error('Faltan hojas obligatorias: Personajes y/o Habilidades')
      const chars = parseCharacters(wsP)
      const abils = parseAbilities(wsH)
      const wsL = wb.Sheets['Listas']
      if (wsL) applyListas(wsL, abils)

      if (!Object.keys(chars).length) throw new Error('No se encontraron filas válidas en Personajes')
      if (!abils.length) throw new Error('No se encontraron filas válidas en Habilidades')

      const ds: Dataset = { characters: chars, abilities: abils }
      setData(ds); setTurns([]); setEffects([]); setHeals([]); setAbilityName('')
      setMessage(`Cargado: ${Object.keys(ds.characters).length} personajes, ${ds.abilities.length} habilidades.`)
    } catch (err:any) {
      setError(err?.message || 'Error leyendo Excel')
    } finally {
      e.target.value = ''
    }
  }

  const bumpHP = (name: string, delta: number) => setHpAdjust(prev => ({ ...prev, [name]: (prev[name] ?? 0) + delta }))
  const setHPAdj = (name: string, value: number) => setHpAdjust(prev => ({ ...prev, [name]: value }))
  const resetHPAdj = (name: string) => setHpAdjust(prev => ({ ...prev, [name]: 0 }))

  const toggleSelected = (name: string) => setSelected(prev => ({ ...prev, [name]: !prev[name] }))
  const selectAll = () => setSelected(Object.fromEntries(names.map(n => [n, true])) as Record<string, boolean>)
  const clearSel = () => setSelected(Object.fromEntries(names.map(n => [n, false])) as Record<string, boolean>)

  function effectsFor(name: string) { return effects.filter(e => e.owner === name) }

  return (
    <div className="container">
      <header className="flex">
        <h1>Calculadora de combate (Excel import, sin BD)</h1>
        <div className="space-x">
          <button className="btn" onClick={undoLast}>Deshacer último</button>
          <button className="btn" onClick={resetAll}>Reiniciar</button>
          <button className="btn" onClick={clearAllEffects}>Quitar todos los efectos</button>
        </div>
      </header>

      <div className="card" style={{marginTop:16}}>
        <div className="row row-2">
          <div>
            <label>Subir Excel (.xlsx) — Hojas: Personajes, Habilidades (y opcional Listas)</label>
            <input type="file" accept=".xlsx,.xls" onChange={onExcelPicked} />
            {message && <div className="success" style={{marginTop:6}}>{message}</div>}
            {error && <div className="error" style={{marginTop:6}}>{error}</div>}
          </div>
          <div>
            <div className="muted">Formato esperado:</div>
            <ul className="muted" style={{marginTop:6}}>
              <li><b>Personajes</b>: Nombre, Vida, Fuerza, Def. Física, Def. Mágica, Magia, Velocidad, <b>Nivel de riesgo</b></li>
              <li><b>Habilidades</b>: Habilidad, Tipo(Físico|Mágico|Buff), Hace Daño (1/0), Bonus Daño, Objetivo(Atacante|Defensor), Mod Fuerza, Mod Def. Física, Mod Def. Mágica, Mod Magia, Mod Velocidad, <b>Mod RL (negativo permitido)</b>, <b>Mod HP</b>, <b>HP Tipo</b> (Barrera|Curación), (opcional) AllowedFor, Duración (turnos)</li>
              <li><b>Listas</b> (opcional): Personaje, Habilidad — una fila por relación</li>
            </ul>
          </div>
        </div>
      </div>

      <div className="card" style={{marginTop:16}}>
        <h2>Participantes en combate</h2>
        <div className="controls-inline" style={{marginTop:8}}>
          <button className="btn" onClick={selectAll}>Seleccionar todos</button>
          <button className="btn" onClick={clearSel}>Limpiar</button>
          <span className="small">Solo los seleccionados aparecen en Orden de turnos, Fichas y desplegables.</span>
        </div>
        <div className="checkbox-grid" style={{marginTop:12}}>
          {names.map(n => (
            <label key={n} style={{display:'flex', alignItems:'center', gap:8}}>
              <input type="checkbox" checked={!!selected[n]} onChange={()=>toggleSelected(n)} />
              {n}
            </label>
          ))}
        </div>
      </div>

      <div className="card" style={{marginTop:16}}>
        <div className="row row-2">
          <div>
            <label>Atacante</label>
            <select value={attacker} onChange={e=>setAttacker(e.target.value)}>
              {selectedNames.map(n => <option key={n} value={n}>{n}</option>)}
            </select>
          </div>
          <div>
            <label>Defensor</label>
            <select value={defender} onChange={e=>setDefender(e.target.value)}>
              {selectedNames.map(n => <option key={n} value={n}>{n}</option>)}
            </select>
          </div>
        </div>
        <div className="row" style={{marginTop:12}}>
          <div>
            <label>Habilidad</label>
            <select value={abilityName} onChange={e=>setAbilityName(e.target.value)}>
              <option value="">— Selecciona —</option>
              {allowedAbilities.map(a => (
                <option key={a.name} value={a.name}>
                  {a.name} ({a.type}{a.doesDamage ? '' : ' • buff'}{(a.duration ?? null) !== null ? ` • ${a.duration}t` : ''})
                </option>
              ))}
            </select>
          </div>
          <div style={{alignSelf:'end'}}>
            <button className="btn primary" onClick={computeAndAddTurn} disabled={!attacker || !defender || !abilityName}>
              Agregar acción
            </button>
          </div>
        </div>
      </div>

      <div className="card" style={{marginTop:16}}>
        <h2>Acción custom (manual)</h2>
        <div className="row row-4" style={{marginTop:8}}>
          <div>
            <label>Nombre</label>
            <input value={cName} onChange={e=>setCName(e.target.value)} />
          </div>
          <div>
            <label>Hace daño</label>
            <select value={cDoesDamage ? '1':'0'} onChange={e=>setCDoesDamage(e.target.value==='1')}>
              <option value="1">Sí</option>
              <option value="0">No</option>
            </select>
          </div>
          <div>
            <label>Tipo de daño</label>
            <select value={cType} onChange={e=>setCType(e.target.value as any)}>
              <option>Físico</option>
              <option>Mágico</option>
              <option>Buff</option>
            </select>
          </div>
          <div>
            <label>Objetivo de modificadores</label>
            <select value={cTarget} onChange={e=>setCTarget(e.target.value as any)}>
              <option>Atacante</option>
              <option>Defensor</option>
            </select>
          </div>
        </div>

        <div className="row row-4" style={{marginTop:8}}>
          <div>
            <label>Daño extra</label>
            <input type="number" value={cExtraDamage} onChange={e=>setCExtraDamage(Number(e.target.value||0))} />
          </div>
          <div>
            <label>HP Tipo</label>
            <select value={cHPMode} onChange={e=>setCHPMode(e.target.value as any)}>
              <option value="">—</option>
              <option value="Barrera">Barrera</option>
              <option value="Curación">Curación</option>
            </select>
          </div>
          <div>
            <label>HP Cantidad</label>
            <input type="number" value={cHP} onChange={e=>setCHP(Number(e.target.value||0))} />
          </div>
          <div />
        </div>

        <div className="row row-4" style={{marginTop:8}}>
          <div>
            <label>Mod Fuerza</label>
            <input type="number" value={cMods.str||0} onChange={e=>setCMods(m=>({ ...m, str: Number(e.target.value||0) }))} />
          </div>
          <div>
            <label>Mod Def. Física</label>
            <input type="number" value={cMods.defFis||0} onChange={e=>setCMods(m=>({ ...m, defFis: Number(e.target.value||0) }))} />
          </div>
          <div>
            <label>Mod Def. Mágica</label>
            <input type="number" value={cMods.defMag||0} onChange={e=>setCMods(m=>({ ...m, defMag: Number(e.target.value||0) }))} />
          </div>
          <div>
            <label>Mod Magia</label>
            <input type="number" value={cMods.mag||0} onChange={e=>setCMods(m=>({ ...m, mag: Number(e.target.value||0) }))} />
          </div>
        </div>

        <div className="row row-4" style={{marginTop:8}}>
          <div>
            <label>Mod Velocidad</label>
            <input type="number" value={cMods.spd||0} onChange={e=>setCMods(m=>({ ...m, spd: Number(e.target.value||0) }))} />
          </div>
          <div>
            <label>Mod RL (puede ser negativo)</label>
            <input type="number" value={cMods.risk||0} onChange={e=>setCMods(m=>({ ...m, risk: Number(e.target.value||0) }))} />
          </div>
          <div>
            <label>Duración (turnos, vacío = ∞)</label>
            <input type="number" value={cDur ?? ''} onChange={e=>{
              const v = e.target.value
              setCDur(v === '' ? null : Number(v))
            }} />
          </div>
          <div style={{alignSelf:'end'}}>
            <button className="btn primary" onClick={computeCustomTurn} disabled={!attacker || !defender}>
              Agregar acción custom
            </button>
          </div>
        </div>
      </div>

      <div className="card" style={{marginTop:16}}>
        <h2>Orden de turnos</h2>
        <div className="order-list" style={{marginTop:8}}>
          {selectedNames.length === 0 ? (
            <div className="muted">Sin participantes seleccionados.</div>
          ) : turnOrder.map((p, idx) => (
            <div className="order-item" key={p.name}>
              <div><b>{idx+1}.</b> {p.name}</div>
              <div className="badge">Velocidad: {p.spd}</div>
            </div>
          ))}
        </div>
        <div className="muted" style={{marginTop:6}}>* Dinámico: si cambia la Velocidad efectiva, cambia el orden.</div>
      </div>

      <div className="card" style={{marginTop:16}}>
        <h2>Registro de combate</h2>
        <div style={{overflowX:'auto', marginTop:8}}>
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Atacante</th>
                <th>Defensor</th>
                <th>Habilidad</th>
                <th>Daño base</th>
                <th>Tipo</th>
                <th>Defensa aplicada</th>
                <th>Daño + hab.</th>
                <th>Daño final</th>
                <th>Vida defensor</th>
              </tr>
            </thead>
            <tbody>
              {turns.length === 0 ? (
                <tr><td colSpan={10} className="muted">Sin acciones todavía.</td></tr>
              ) : (
                turns.map((t,i) => {
                  const lifeAfter = currentHP(t.defender, i)
                  return (
                    <tr key={i}>
                      <td>{i+1}</td>
                      <td>{t.attacker}</td>
                      <td>{t.defender}</td>
                      <td>{t.abilityName}</td>
                      <td>{t.baseDamage ?? '—'}</td>
                      <td>{t.type}</td>
                      <td>{t.defenseApplied ?? '—'}</td>
                      <td>{t.damagePlusAbility ?? '—'}</td>
                      <td><b>{t.finalDamage}</b></td>
                      <td>{lifeAfter}</td>
                    </tr>
                  )
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div className="row row-3" style={{marginTop: 16}}>
        {selectedNames.map(name => {
          const mods = currentModsFor(name)
          const eff = {
            str: effectiveStat(name, 'str'),
            defFis: effectiveStat(name, 'defFis'),
            defMag: effectiveStat(name, 'defMag'),
            mag: effectiveStat(name, 'mag'),
            spd: effectiveStat(name, 'spd'),
            risk: effectiveStat(name, 'risk'),
          }
          const hp = turns.length ? currentHP(name, turns.length - 1) : (data.characters[name].vida + (hpAdjust[name] ?? 0))
          const efList = effects.filter(e => e.owner === name)
          const barrier = totalBarrier(name)
          const labelMap: Record<string,string> = { str:'Fuerza', defFis:'DefFis', defMag:'DefMag', mag:'Magia', spd:'Vel', risk:'RL' }
          return (
            <div className="card" key={name}>
              <div className="flex">
                <h3>{name}</h3>
                <div className="pill">turnos: {turns.length}</div>
              </div>
              <div style={{marginTop:8}}>
                <table>
                  <tbody>
                    <tr><td><b>Vida actual</b></td><td>{hp} {barrier>0 ? <span className="badge">Barrera: {barrier}</span> : null}</td></tr>
                    <tr>
                      <td>Ajuste de vida (±)</td>
                      <td>
                        <div className="controls-inline">
                          <button className="btn" onClick={()=>bumpHP(name,-10)}>-10</button>
                          <button className="btn" onClick={()=>bumpHP(name,-1)}>-1</button>
                          <input type="number" value={hpAdjust[name] ?? 0} onChange={e=>setHPAdj(name, Number(e.target.value || 0))} />
                          <button className="btn" onClick={()=>bumpHP(name,+1)}>+1</button>
                          <button className="btn" onClick={()=>bumpHP(name,+10)}>+10</button>
                          <button className="btn" onClick={()=>resetHPAdj(name)}>Reset</button>
                        </div>
                        <div className="small">Este ajuste se suma/resta a la vida total tras daños del registro.</div>
                      </td>
                    </tr>
                    <tr><td>Fuerza (ef.)</td><td>{eff.str} {mods.str ? <span className="muted">(mod {mods.str! >= 0 ? '+' : ''}{mods.str})</span> : null}</td></tr>
                    <tr><td>Def. Física (ef.)</td><td>{eff.defFis} {mods.defFis ? <span className="muted">(mod {mods.defFis! >= 0 ? '+' : ''}{mods.defFis})</span> : null}</td></tr>
                    <tr><td>Def. Mágica (ef.)</td><td>{eff.defMag} {mods.defMag ? <span className="muted">(mod {mods.defMag! >= 0 ? '+' : ''}{mods.defMag})</span> : null}</td></tr>
                    <tr><td>Magia (ef.)</td><td>{eff.mag} {mods.mag ? <span className="muted">(mod {mods.mag! >= 0 ? '+' : ''}{mods.mag})</span> : null}</td></tr>
                    <tr><td>Velocidad (ef.)</td><td>{eff.spd} {mods.spd ? <span className="muted">(mod {mods.spd! >= 0 ? '+' : ''}{mods.spd})</span> : null}</td></tr>
                    <tr><td>Nivel de riesgo (ef.)</td><td>{eff.risk} {typeof mods.risk === 'number' ? <span className="muted">(mod {mods.risk >= 0 ? '+' : ''}{mods.risk})</span> : null}</td></tr>
                  </tbody>
                </table>
              </div>

              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center', marginTop:10}}>
                <h3>Efectos activos</h3>
                {efList.length > 0 && (
                  <button className="btn" onClick={()=>clearEffectsFor(name)}>
                    Quitar todos
                  </button>
                )}
              </div>

              {efList.length === 0 ? (
                <div className="muted">Sin efectos.</div>
              ) : (
                <div className="effect-list">
                  {efList.map(ef => (
                    <div className="effect-row" key={ef.id}>
                      <div>
                        <span className="effect-name">{ef.abilityName}</span>{' '}
                        <span className="muted">
                          {ef.remaining === null ? '∞' : `${ef.remaining}t`} ·
                          {Object.entries(ef.mods).map(([k,v]) => (v && v!==0) ? ` ${labelMap[k] ?? k} ${v>=0?'+':''}${v}` : '').join('') || ' sin mods'}
                          {(ef.barrierRemaining ?? 0) > 0 ? ` · Barrera ${ef.barrierRemaining}/${ef.barrierMax ?? ef.barrierRemaining}` : ''}
                        </span>
                      </div>
                      <div className="effect-controls">
                        {ef.remaining !== null && (<>
                          <button className="btn" onClick={()=>setEffects(prev=>prev.map(e=>e.id===ef.id?{...e, remaining: Math.max(0,(e.remaining??0)-1)}:e))}>-1</button>
                          <input type="number" value={ef.remaining ?? 0} onChange={e=>setEffects(prev=>prev.map(x=>x.id===ef.id?{...x, remaining: Math.max(0, Number(e.target.value||0))}:x))} />
                          <button className="btn" onClick={()=>setEffects(prev=>prev.map(e=>e.id===ef.id?{...e, remaining: (e.remaining??0)+1}:e))}>+1</button>
                        </>)}
                        {(ef.barrierRemaining ?? null) !== null && (<>
                          <button className="btn" onClick={()=>setEffects(prev=>prev.map(e=>e.id===ef.id?{...e, barrierRemaining: Math.max(0,(e.barrierRemaining??0)-10)}:e))}>-10</button>
                          <button className="btn" onClick={()=>setEffects(prev=>prev.map(e=>e.id===ef.id?{...e, barrierRemaining: Math.max(0,(e.barrierRemaining??0)-1)}:e))}>-1</button>
                          <input type="number" value={ef.barrierRemaining ?? 0} onChange={e=>setEffects(prev=>prev.map(x=>x.id===ef.id?{...x, barrierRemaining: Math.max(0, Number(e.target.value||0))}:x))} />
                          <button className="btn" onClick={()=>setEffects(prev=>prev.map(e=>e.id===ef.id?{...e, barrierRemaining: (e.barrierRemaining??0)+1}:e))}>+1</button>
                          <button className="btn" onClick={()=>setEffects(prev=>prev.map(e=>e.id===ef.id?{...e, barrierRemaining: (e.barrierRemaining??0)+10}:e))}>+10</button>
                        </>)}
                        <button className="btn" onClick={()=>setEffects(prev=>prev.filter(e=>e.id!==ef.id))}>Quitar</button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )
        })}
      </div>

      <div className="muted" style={{marginTop:8}}>
        * Las barreras dan bonos mientras su vida &gt; 0. Curaciones son instantáneas. RL puede ser negativo por debuffs.
      </div>
    </div>
  )
}
