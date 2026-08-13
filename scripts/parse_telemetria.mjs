// Lógica de leitura do "Relatório de Eventos" bruto exportado da MOVIAS.
//
// Isto é um espelho, linha a linha, da função `parseTelemetriaWorkbook` (e das
// funções auxiliares que ela usa) que existe dentro de index.html, no bloco
// "Atualizar dados (Telemetria)". As duas cópias precisam ficar em sincronia:
// se um dia mexer numa regra de parsing (ex: novo teto de plausibilidade do
// odômetro, novo tipo de evento de alerta), replique a mudança na outra cópia.
// Motivo de não ter uma cópia só: o index.html roda 100% no navegador (sem
// build step, arquivo único) e este script roda em Node dentro do GitHub
// Actions — não dá pra importar um dentro do outro sem montar um bundler.

import XLSX from 'xlsx';

export const FLEET_TOTAL_CACAMBAS = 10;
export const PCT_MANUAL_IDENTIFICADO = 2.6;
export const REQUIRED_COLS = ['Frota', 'Mapa', 'Descrição do Evento', 'Peso', 'Data Evento'];

export function haversineKm(lat1, lng1, lat2, lng2) {
  if (lat1 == null || lng1 == null || lat2 == null || lng2 == null) return null;
  const R = 6371.0088;
  const toRad = d => d * Math.PI / 180;
  const p1 = toRad(lat1), p2 = toRad(lat2);
  const dphi = toRad(lat2 - lat1), dlmb = toRad(lng2 - lng1);
  const a = Math.sin(dphi / 2) ** 2 + Math.cos(p1) * Math.cos(p2) * Math.sin(dlmb / 2) ** 2;
  return R * 2 * Math.asin(Math.sqrt(a));
}

export function cleanLabel(txt) {
  if (!txt) return txt;
  let t = String(txt).trim();
  t = t.replace(/,\s*São Paulo,\s*\d{5}-\d{3},\s*Brasil\s*$/, '');
  t = t.replace(/,\s*São Paulo,\s*Brasil\s*$/, '');
  t = t.replace(/,\s*Brasil\s*$/, '');
  return t.trim();
}

export function parsePeso(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim().replace(',', '.');
  if (s === '') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

// Odômetro (km) do veículo no momento do evento — coluna nova da MOVIAS (nem todo
// evento tem leitura; ficou sem sensor de odômetro habilitado antes de determinada
// data/veículo).
export function parseOdometro(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim().replace(',', '.');
  if (s === '') return null;
  const n = parseFloat(s);
  return (isNaN(n) || n <= 0) ? null : n;
}

export function parseDataEvento(v) {
  if (v instanceof Date) return v;
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})/);
  if (!m) return null;
  const [, dd, mm, yyyy, HH, MM, SS] = m.map(Number);
  return new Date(yyyy, mm - 1, dd, HH, MM, SS);
}

export function toISODate(d) {
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

export function findHeaderRow(ws) {
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let r = range.s.r; r <= Math.min(range.e.r, 20); r++) {
    const rowVals = {};
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') rowVals[String(cell.v).trim()] = c;
    }
    if (REQUIRED_COLS.every(req => req in rowVals)) return { row: r, cols: rowVals };
  }
  return null;
}

// Recalcula o bloco `meta` a partir de uma lista de viagens (usado tanto logo após
// o parsing de um único relatório quanto — no robô de automação — depois de fundir
// o relatório novo com o histórico acumulado, já que o relatório "Diário" da MOVIAS
// só traz uma janela móvel de ~48h, não o histórico completo).
export function buildMeta(trips) {
  const frotaTipo = {};
  trips.forEach(t => { if (t.TipoVeiculo) frotaTipo[t.Frota] = t.TipoVeiculo; });
  const tiposEquipamento = [...new Set(Object.values(frotaTipo))].sort();
  const equipPorTipo = {};
  Object.entries(frotaTipo).forEach(([f, t]) => { (equipPorTipo[t] = equipPorTipo[t] || []).push(f); });
  Object.keys(equipPorTipo).forEach(k => { equipPorTipo[k] = [...new Set(equipPorTipo[k])].sort(); });
  const todasFrotas = Object.keys(frotaTipo).sort();

  const diasPeriodo = [...new Set(trips.map(t => t.Data))].sort();
  const cacambasMonitoradas = new Set(trips.map(t => t.Frota)).size;
  const viagensDiaMedia = cacambasMonitoradas && diasPeriodo.length ? Math.round((trips.length / (cacambasMonitoradas * diasPeriodo.length)) * 10) / 10 : 0;

  return {
    periodo_min: diasPeriodo[0], periodo_max: diasPeriodo[diasPeriodo.length - 1], dias_periodo: diasPeriodo.length,
    frota_total_cacambas: FLEET_TOTAL_CACAMBAS, pct_manual_identificado: PCT_MANUAL_IDENTIFICADO,
    cacambas_monitoradas_piloto: cacambasMonitoradas, viagens_dia_media_por_cacamba: viagensDiaMedia,
    tipos_equipamento: tiposEquipamento, equip_por_tipo: equipPorTipo, todas_frotas: todasFrotas, frota_tipo: frotaTipo,
  };
}

export function parseTelemetriaWorkbook(wb) {
  let ws = wb.Sheets['Relatório de Eventos'];
  let found = ws ? findHeaderRow(ws) : null;
  if (!found) {
    for (const name of wb.SheetNames) {
      const cand = findHeaderRow(wb.Sheets[name]);
      if (cand) { ws = wb.Sheets[name]; found = cand; break; }
    }
  }
  if (!found) {
    throw new Error(`Colunas essenciais não encontradas (${REQUIRED_COLS.join(', ')}). Verifique se é o relatório bruto exportado da MOVIAS (aba "Relatório de Eventos"), sem tratamento manual.`);
  }
  const headerRow = found.row, cols = found.cols;
  const hasLatLong = 'Lat/Long' in cols;
  const hasOdometro = 'Odômetro' in cols;
  const range = XLSX.utils.decode_range(ws['!ref']);
  const getCell = (r, colName) => (colName in cols) ? ws[XLSX.utils.encode_cell({ r, c: cols[colName] })] : undefined;

  const rows = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const frotaCell = getCell(r, 'Frota');
    const frota = frotaCell && frotaCell.v;
    if (!frota) continue;
    const mapaCell = getCell(r, 'Mapa');
    const mapaTxt = mapaCell ? mapaCell.v : null;

    let lat = null, lng = null;
    if (hasLatLong) {
      const llCell = getCell(r, 'Lat/Long');
      if (llCell && llCell.v) {
        const m = String(llCell.v).match(/^\s*(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)\s*$/);
        if (m) { lat = parseFloat(m[1]); lng = parseFloat(m[2]); }
      }
    }
    if (lat === null && mapaCell && mapaCell.l && mapaCell.l.Target) {
      const m = mapaCell.l.Target.match(/q=(-?\d+\.\d+),(-?\d+\.\d+)/);
      if (m) { lat = parseFloat(m[1]); lng = parseFloat(m[2]); }
    }

    const dataCell = getCell(r, 'Data Evento');
    const dataEvt = dataCell ? parseDataEvento(dataCell.v) : null;
    if (!dataEvt) continue;

    const pesoCell = getCell(r, 'Peso');
    const tipoCell = getCell(r, 'Tipo de Veículo');
    const placaCell = getCell(r, 'Placa');
    const eventoCell = getCell(r, 'Descrição do Evento');
    const odoCell = hasOdometro ? getCell(r, 'Odômetro') : undefined;

    rows.push({
      Placa: placaCell ? placaCell.v : null,
      Frota: String(frota).trim(),
      TipoVeiculo: tipoCell ? tipoCell.v : null,
      DataEvento: dataEvt,
      MapaLabel: cleanLabel(mapaTxt),
      Lat: lat, Lng: lng,
      Evento: eventoCell ? eventoCell.v : null,
      Peso: pesoCell ? parsePeso(pesoCell.v) : null,
      Odometro: odoCell ? parseOdometro(odoCell.v) : null,
    });
  }
  if (rows.length === 0) {
    throw new Error('Nenhuma linha de evento válida foi encontrada (verifique se a coluna "Frota" está preenchida e "Data Evento" segue o formato dd/mm/aaaa hh:mm:ss).');
  }

  const byFrota = {};
  rows.forEach(r => { (byFrota[r.Frota] = byFrota[r.Frota] || []).push(r); });
  Object.values(byFrota).forEach(list => list.sort((a, b) => a.DataEvento - b.DataEvento));

  const trips = [];
  const alertEvents = [];
  const ALERT_TYPES = ['Desabastecimento', 'Bateria Principal Desconectada'];

  Object.entries(byFrota).forEach(([frota, evs]) => {
    let lastCarregado = null;
    evs.forEach(e => {
      if (ALERT_TYPES.includes(e.Evento)) {
        alertEvents.push({ Frota: frota, TipoVeiculo: e.TipoVeiculo, Evento: e.Evento, DataEvento: e.DataEvento.toISOString(), MapaLabel: e.MapaLabel, Lat: e.Lat, Lng: e.Lng });
        return;
      }
      if (e.Evento === 'Sensor Carregado') {
        lastCarregado = e;
      } else if (e.Evento === 'Sensor Descarregando') {
        const origem = lastCarregado;
        const distLinhaReta = origem ? haversineKm(origem.Lat, origem.Lng, e.Lat, e.Lng) : null;
        const gapCargaDescargaMin = origem ? (e.DataEvento - origem.DataEvento) / 60000 : null;
        let distKm = distLinhaReta, distFonte = distLinhaReta !== null ? 'linha_reta' : null;
        if (origem && origem.Odometro !== null && e.Odometro !== null && gapCargaDescargaMin <= 180) {
          const distOdo = e.Odometro - origem.Odometro;
          const plausivel = distOdo > 0 && distOdo < 200 && (distLinhaReta === null || distOdo >= distLinhaReta * 0.85);
          if (plausivel) { distKm = distOdo; distFonte = 'odometro'; }
        }
        trips.push({
          Frota: frota, TipoVeiculo: e.TipoVeiculo, Placa: e.Placa,
          DataCarregado: origem ? origem.DataEvento.toISOString() : null,
          DataDescarga: e.DataEvento.toISOString(),
          Data: toISODate(e.DataEvento),
          OrigemLabel: origem ? origem.MapaLabel : 'Origem fora do período coletado',
          OrigemLat: origem ? origem.Lat : null, OrigemLng: origem ? origem.Lng : null,
          DestinoLabel: e.MapaLabel, DestinoLat: e.Lat, DestinoLng: e.Lng,
          PesoTon: e.Peso,
          DistanciaKm: distKm !== null ? Math.round(distKm * 1000) / 1000 : null,
          DistanciaFonte: distFonte,
          MomentoTonKm: (distKm !== null && e.Peso !== null) ? Math.round(e.Peso * distKm * 100) / 100 : null,
          DuracaoMin: origem ? Math.round((e.DataEvento - origem.DataEvento) / 6000) / 10 : null,
          OrigemIdentificada: !!origem,
        });
      }
    });
  });
  if (trips.length === 0) {
    throw new Error('Nenhuma viagem foi identificada (não há eventos "Sensor Descarregando" no arquivo). Confirme se o período exportado inclui caçambas com sensor de báscula ativo.');
  }

  const meta = buildMeta(trips);
  return { meta, trips, alertEvents };
}
