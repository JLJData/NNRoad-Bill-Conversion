#!/usr/bin/env node
/**
 * HyperFormula 公式快照（模板三引擎压测用）
 *
 * 用法:
 *   node hf_snapshot.mjs <xlsx路径> [--sheet PN] [--max 300]
 *
 * 依赖：从 HRONE_OFFICE_UI（或默认 ../GIT 其他项目/hrone-office-ui）加载 luckyexcel + hyperformula
 * 输出 JSON 到 stdout（与 excel-snapshot 同形）
 */
import fs from 'fs'
import path from 'path'
import module from 'module'
import { fileURLToPath } from 'url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

function resolveOfficeUi() {
  if (process.env.HRONE_OFFICE_UI) return path.resolve(process.env.HRONE_OFFICE_UI)
  const candidates = [
    path.resolve(__dirname, '..', 'GIT 其他项目', 'hrone-office-ui'),
    path.resolve(__dirname, '..', 'hrone-office-ui'),
  ]
  for (const c of candidates) {
    if (fs.existsSync(path.join(c, 'node_modules', 'hyperformula'))) return c
  }
  return candidates[0]
}

const OFFICE_UI = resolveOfficeUi()
const requireFromUi = module.createRequire(path.join(OFFICE_UI, 'package.json'))

// 抑制 LuckyExcel 往 stdout 打的噪音，避免污染 JSON
const _log = console.log.bind(console)
const _info = console.info.bind(console)
console.log = (...args) => {
  if (args.length === 1 && typeof args[0] === 'string' && args[0].startsWith('{')) _log(...args)
}
console.info = () => {}

let LuckyExcel
let HyperFormula
try {
  LuckyExcel = requireFromUi('luckyexcel')
  ;({ HyperFormula } = requireFromUi('hyperformula'))
} catch (e) {
  _log(
    JSON.stringify({
      ok: false,
      engine: 'hyperformula',
      sheetName: '',
      cells: [],
      truncated: false,
      message: `无法加载 hyperformula/luckyexcel（OFFICE_UI=${OFFICE_UI}）: ${e.message}`,
    }),
  )
  process.exit(2)
}

function materialize(sheet) {
  const celldata = sheet.celldata
  if (!Array.isArray(celldata) || !celldata.length) return
  const hasData =
    Array.isArray(sheet.data) &&
    sheet.data.some((row) => Array.isArray(row) && row.some((c) => c != null))
  if (hasData) return
  let maxR = 0
  let maxC = 0
  for (const item of celldata) {
    const r = Number(item?.r)
    const c = Number(item?.c)
    if (!Number.isFinite(r) || !Number.isFinite(c)) continue
    if (r > maxR) maxR = r
    if (c > maxC) maxC = c
  }
  const data = []
  for (let r = 0; r <= maxR; r++) data[r] = new Array(maxC + 1).fill(null)
  for (const item of celldata) {
    const r = Number(item?.r)
    const c = Number(item?.c)
    if (!Number.isFinite(r) || !Number.isFinite(c)) continue
    const v = item?.v
    data[r][c] = v != null && typeof v === 'object' ? { ...v } : v
  }
  sheet.data = data
}

function normalizeFormula(f) {
  let s = String(f || '').trim()
  if (!s) return ''
  s = s.replace(/<!\[CDATA\[/gi, '').replace(/\]\]>/g, '').trim()
  if (!s.startsWith('=')) s = `=${s}`
  s = s.replace(/^=\s*\+/, '=')
  s = s.replace(/&\s*\+/g, '&')
  s = s.replace(/(^|[^A-Za-z0-9_.])(\d+(?:\.\d+)?)\s*%/g, (_, a, n) => `${a}(${n}/100)`)
  return s
}

function isUnsafe(formula) {
  const f = String(formula || '')
  if (/EOMONTH\s*\(/i.test(f)) return true
  if (/\b(INDIRECT|OFFSET|CELL)\s*\(/i.test(f)) return true
  return false
}

function toRaw(cell, frozen) {
  if (cell == null) return null
  if (typeof cell !== 'object') {
    if (typeof cell === 'number' && Number.isFinite(cell)) return cell
    if (typeof cell === 'boolean') return cell
    if (typeof cell === 'string' && cell !== '') {
      const n = Number(String(cell).replace(/,/g, ''))
      return Number.isFinite(n) ? n : cell
    }
    return null
  }
  if (typeof cell.f === 'string' && cell.f.trim()) {
    const f = normalizeFormula(cell.f)
    if (f && !isUnsafe(f)) return f
    frozen.add(true)
    if (typeof cell.v === 'number' && Number.isFinite(cell.v)) return cell.v
    const t = String(cell.m ?? cell.v ?? '').replace(/[$,\s,，]/g, '')
    const n = Number(t)
    return Number.isFinite(n) ? n : t || null
  }
  if (typeof cell.v === 'number' && Number.isFinite(cell.v)) return cell.v
  if (typeof cell.v === 'boolean') return cell.v
  if (typeof cell.v === 'string' && cell.v !== '') {
    const n = Number(String(cell.v).replace(/,/g, ''))
    return Number.isFinite(n) ? n : cell.v
  }
  return null
}

function colLetter(col0) {
  let n = col0 + 1
  let s = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    n = Math.floor((n - 1) / 26)
  }
  return s || 'A'
}

function parseArgs(argv) {
  const out = { path: '', sheet: 'PN', max: 300 }
  const rest = [...argv]
  out.path = rest.shift() || ''
  while (rest.length) {
    const a = rest.shift()
    if (a === '--sheet') out.sheet = rest.shift() || 'PN'
    else if (a === '--max') out.max = Math.max(1, Number(rest.shift() || 300) || 300)
  }
  if (out.sheet === '') out.sheet = null
  return out
}

function snapshot(filePath, sheetFilter, maxCells) {
  return new Promise((resolve) => {
    let buf
    try {
      buf = fs.readFileSync(filePath)
    } catch (e) {
      resolve({
        ok: false,
        engine: 'hyperformula',
        sheetName: sheetFilter || '',
        cells: [],
        truncated: false,
        message: `读文件失败: ${e.message}`,
      })
      return
    }
    LuckyExcel.transformExcelToLucky(buf, (json) => {
      try {
        const sheets = json?.sheets || []
        if (!sheets.length) {
          resolve({
            ok: false,
            engine: 'hyperformula',
            sheetName: sheetFilter || '',
            cells: [],
            truncated: false,
            message: 'LuckyExcel 无 sheets',
          })
          return
        }
        sheets.forEach(materialize)
        const init = {}
        for (const s of sheets) {
          const name = String(s.name || '').trim()
          if (!name) continue
          const data = s.data || [[]]
          let maxC = 1
          for (const row of data) if (Array.isArray(row) && row.length > maxC) maxC = row.length
          const matrix = []
          for (let r = 0; r < data.length; r++) {
            const line = new Array(maxC).fill(null)
            const frozen = new Set()
            for (let c = 0; c < maxC; c++) line[c] = toRaw(data[r]?.[c], frozen)
            matrix.push(line)
          }
          init[name] = matrix
        }
        const hf = HyperFormula.buildFromSheets(init, {
          licenseKey: 'gpl-v3',
          useColumnIndex: false,
          evaluateNullToZero: true,
        })
        const want = (sheetFilter || '').trim().toLowerCase()
        let targetName = ''
        for (const name of Object.keys(init)) {
          if (!want || name.toLowerCase() === want) {
            targetName = name
            break
          }
        }
        if (!targetName) {
          hf.destroy()
          resolve({
            ok: false,
            engine: 'hyperformula',
            sheetName: sheetFilter || '',
            cells: [],
            truncated: false,
            message: `未找到工作表: ${sheetFilter}`,
          })
          return
        }
        const sheetId = hf.getSheetId(targetName)
        const sh = sheets.find((s) => String(s.name) === targetName)
        const data = sh?.data || []
        const cells = []
        let truncated = false
        for (let r = 0; r < data.length; r++) {
          const row = data[r]
          if (!row) continue
          for (let c = 0; c < row.length; c++) {
            const cell = row[c]
            if (!cell || typeof cell !== 'object') continue
            if (typeof cell.f !== 'string' || !cell.f.trim()) continue
            if (cells.length >= maxCells) {
              truncated = true
              break
            }
            let raw
            try {
              raw = hf.getCellValue({ sheet: sheetId, row: r, col: c })
            } catch {
              continue
            }
            if (raw == null || (typeof raw === 'object' && raw !== null)) {
              // DetailedCellError or empty — skip numeric-less
              if (typeof raw === 'object' && raw != null) continue
              continue
            }
            let numberValue = null
            let textValue = ''
            if (typeof raw === 'number' && Number.isFinite(raw)) {
              numberValue = raw
              textValue = String(raw)
            } else if (typeof raw === 'boolean') {
              textValue = raw ? 'TRUE' : 'FALSE'
            } else if (typeof raw === 'string') {
              textValue = raw
              const n = Number(String(raw).replace(/[$,\s,，%]/g, ''))
              if (Number.isFinite(n)) numberValue = n
            }
            if (numberValue == null && !textValue) continue
            cells.push({
              sheet: targetName,
              ref: `${colLetter(c)}${r + 1}`,
              row: r,
              col: c,
              formula: normalizeFormula(cell.f),
              numberValue,
              textValue,
            })
          }
          if (truncated) break
        }
        hf.destroy()
        resolve({
          ok: true,
          engine: 'hyperformula',
          sheetName: targetName,
          cells,
          truncated,
          message: truncated ? `公式格较多，仅抽取前 ${maxCells} 个` : '',
          cellCount: cells.length,
        })
      } catch (e) {
        resolve({
          ok: false,
          engine: 'hyperformula',
          sheetName: sheetFilter || '',
          cells: [],
          truncated: false,
          message: `HF 快照失败: ${e.message}`,
        })
      }
    })
  })
}

const args = parseArgs(process.argv.slice(2))
if (!args.path) {
  console.log(
    JSON.stringify({
      ok: false,
      engine: 'hyperformula',
      sheetName: '',
      cells: [],
      truncated: false,
      message: '缺少 xlsx 路径',
    }),
  )
  process.exit(2)
}

const result = await snapshot(args.path, args.sheet, args.max)
process.stdout.write(JSON.stringify(result) + '\n')
process.exit(result.ok ? 0 : 2)
