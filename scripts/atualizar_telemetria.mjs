// Robô de atualização da Telemetria: entra na caixa de e-mail dedicada que recebe
// o "Relatório Agendado" diário do MOVIAS (relatório "Evento", em Excel), pega o
// anexo mais recente, reaproveita a mesma lógica de leitura do painel
// (parse_telemetria.mjs — espelho de parseTelemetriaWorkbook do index.html) e
// grava dados_atuais.json na raiz do repositório. O painel busca esse arquivo
// sozinho ao carregar (ver index.html).
//
// Variáveis de ambiente esperadas (definidas como Secrets no GitHub Actions):
//   GMAIL_ADDRESS        e-mail que recebe os relatórios (ex: telemetriasp284@gmail.com)
//   GMAIL_APP_PASSWORD   senha de app gerada nas configurações de segurança do Google
// Opcionais:
//   MOVIAS_SENDER_FILTER  trecho do endereço do remetente a exigir (ex: "movias.com.br").
//                         Vazio = aceita qualquer remetente, desde que o e-mail traga
//                         um anexo .xlsx (ok pra uma caixa dedicada só a isso).
//   LOOKBACK_DAYS         quantos dias pra trás procurar e-mails (padrão: 3)

import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';
import AdmZip from 'adm-zip';
import { parseTelemetriaWorkbook, buildMeta } from './parse_telemetria.mjs';

const GMAIL_ADDRESS = process.env.GMAIL_ADDRESS;
const GMAIL_APP_PASSWORD = process.env.GMAIL_APP_PASSWORD;
const LOOKBACK_DAYS = parseInt(process.env.LOOKBACK_DAYS || '3', 10);
const SENDER_FILTER = (process.env.MOVIAS_SENDER_FILTER || '').trim().toLowerCase();

if (!GMAIL_ADDRESS || !GMAIL_APP_PASSWORD) {
  console.error('Faltam as variáveis de ambiente GMAIL_ADDRESS e/ou GMAIL_APP_PASSWORD (configure como Secrets do repositório).');
  process.exit(1);
}

const OUTPUT_PATH = path.resolve(process.cwd(), 'dados_atuais.json');

function ehAnexoXlsx(a) {
  return /\.xlsx?$/i.test(a.filename || '') ||
    a.contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
}

function ehAnexoZip(a) {
  return /\.zip$/i.test(a.filename || '') ||
    a.contentType === 'application/zip' ||
    a.contentType === 'application/x-zip-compressed' ||
    a.contentType === 'application/octet-stream' && /\.zip$/i.test(a.filename || '');
}

// O "Relatório Agendado" do MOVIAS não vem como anexo de verdade no e-mail — o
// corpo (HTML) traz um botão "Download Zip" que aponta pra um link de download
// no próprio servidor do MOVIAS (válido sem precisar estar logado, já que é
// assim que quem recebe o e-mail consegue baixar o relatório). Por isso, além
// de checar anexos MIME (.xlsx/.xls e .zip, para o caso de um dia mudarem o
// formato do envio), procura também esse link no corpo do e-mail.
function extrairLinkDownloadHtml(html) {
  if (!html) return null;
  const m = html.match(/https:\/\/www\.movias\.com\.br:8443\/api\/report\/download\?group=[0-9a-fA-F-]+/i);
  return m ? m[0] : null;
}

async function baixarDeUrl(url) {
  const resp = await fetch(url, {
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; TripoloniTelemetriaBot/1.0)' },
  });
  if (!resp.ok) {
    throw new Error(`Falha ao baixar o relatório pelo link do MOVIAS (HTTP ${resp.status} ${resp.statusText}): ${url}`);
  }
  const buffer = Buffer.from(await resp.arrayBuffer());
  return buffer;
}

// O relatório "Agrupado" da MOVIAS vem como um .xlsx dentro de um .zip (em vez
// do .xlsx direto). Extrai o primeiro .xlsx encontrado dentro do zip.
function extrairXlsxDoZip(bufferZip, nomeZip) {
  const zip = new AdmZip(bufferZip);
  const entradas = zip.getEntries().filter(e => !e.isDirectory && /\.xlsx?$/i.test(e.entryName));
  if (entradas.length === 0) {
    throw new Error(`O anexo "${nomeZip}" é um .zip mas não contém nenhum .xlsx dentro.`);
  }
  // Se houver mais de um, usa o maior (mais provável de ser o relatório completo,
  // não um resumo/capa).
  entradas.sort((a, b) => b.header.size - a.header.size);
  return { nomeInterno: entradas[0].entryName, buffer: entradas[0].getData() };
}

// Varre a caixa de entrada e retorna TODOS os e-mails (dentro da janela de
// LOOKBACK_DAYS) que trazem um anexo utilizável (.xlsx/.xls direto, ou .zip
// contendo um .xlsx), ordenados do mais antigo pro mais recente. É comum o
// agendamento "Agrupado" do MOVIAS disparar mais de um e-mail por dia (ex.:
// reenvio, ou geração em lotes) — processar todos e deixar a fusão por
// Frota+horário decidir quem fica é mais seguro do que confiar só no mais
// recente, já que a fusão é idempotente (reprocessar o mesmo evento não duplica).
async function encontrarAnexosRelevantes(client) {
  const lock = await client.getMailboxLock('INBOX');
  try {
    const since = new Date();
    since.setDate(since.getDate() - LOOKBACK_DAYS);

    const uids = await client.search({ since }, { uid: true });
    if (!uids || uids.length === 0) {
      console.log(`Nenhum e-mail encontrado nos últimos ${LOOKBACK_DAYS} dia(s).`);
      return [];
    }

    const candidatos = [];
    for await (const msg of client.fetch(uids, { envelope: true, uid: true }, { uid: true })) {
      const fromAddr = (msg.envelope && msg.envelope.from && msg.envelope.from[0] && msg.envelope.from[0].address || '').toLowerCase();
      if (SENDER_FILTER && !fromAddr.includes(SENDER_FILTER)) continue;
      candidatos.push({ uid: msg.uid, date: msg.envelope?.date, subject: msg.envelope?.subject, from: fromAddr });
    }
    candidatos.sort((a, b) => new Date(a.date) - new Date(b.date));

    const encontrados = [];
    let semNadaUtil = 0;
    for (const cand of candidatos) {
      const { content } = await client.download(cand.uid, undefined, { uid: true });
      const parsed = await simpleParser(content);
      const anexoXlsx = (parsed.attachments || []).find(ehAnexoXlsx);
      const anexoZip = (parsed.attachments || []).find(ehAnexoZip);
      const linkZip = extrairLinkDownloadHtml(parsed.html || parsed.textAsHtml || '');
      if (anexoXlsx) {
        encontrados.push({ ...cand, attachment: anexoXlsx, tipo: 'xlsx' });
      } else if (anexoZip) {
        encontrados.push({ ...cand, attachment: anexoZip, tipo: 'zip' });
      } else if (linkZip) {
        encontrados.push({ ...cand, downloadUrl: linkZip, tipo: 'link-zip' });
      } else {
        semNadaUtil++;
      }
    }
    if (semNadaUtil > 0) {
      console.log(`${semNadaUtil} e-mail(s) sem anexo .xlsx/.zip nem link de download do MOVIAS (ignorados).`);
    }
    return encontrados;
  } finally {
    lock.release();
  }
}

async function main() {
  const client = new ImapFlow({
    host: 'imap.gmail.com',
    port: 993,
    secure: true,
    auth: { user: GMAIL_ADDRESS, pass: GMAIL_APP_PASSWORD },
    logger: false,
  });

  await client.connect();
  console.log(`Conectado a ${GMAIL_ADDRESS}.`);

  let encontrados;
  try {
    encontrados = await encontrarAnexosRelevantes(client);
  } finally {
    await client.logout();
  }

  if (!encontrados || encontrados.length === 0) {
    console.log('Nada a atualizar nesta execução.');
    return;
  }

  // O relatório "Diário" da MOVIAS traz só uma janela móvel de ~48h (desde 0h de
  // ontem até 23h59 de hoje) — não o histórico completo do projeto. Por isso o
  // robô nunca substitui dados_atuais.json de uma vez: ele funde o que veio de
  // novo com o que já estava acumulado, usando Frota+horário do evento como chave
  // (o mesmo evento relatado de novo, por causa da sobreposição entre um dia e o
  // próximo, simplesmente cai na mesma chave e não duplica).
  const anterior = fs.existsSync(OUTPUT_PATH) ? JSON.parse(fs.readFileSync(OUTPUT_PATH, 'utf8')) : null;
  const tripsAnteriores = anterior?.trips || [];
  const alertasAnteriores = anterior?.alertEvents || [];

  const tripKey = t => `${t.Frota}|${t.DataDescarga}`;
  const alertKey = a => `${a.Frota}|${a.DataEvento}|${a.Evento}`;

  const tripsMap = new Map(tripsAnteriores.map(t => [tripKey(t), t]));
  const alertasMap = new Map(alertasAnteriores.map(a => [alertKey(a), a]));

  // Processa cada e-mail encontrado (do mais antigo pro mais recente), extraindo
  // o .xlsx de dentro do .zip quando for o caso, e fundindo tudo na mesma base.
  for (const picked of encontrados) {
    let bufferXlsx;
    let nomeUsado;
    if (picked.tipo === 'link-zip') {
      const bufferZip = await baixarDeUrl(picked.downloadUrl);
      if (bufferZip.length < 4 || bufferZip[0] !== 0x50 || bufferZip[1] !== 0x4b) {
        // Não começa com "PK" (assinatura de .zip) — provavelmente voltou uma
        // página de login/erro em vez do arquivo. Loga o começo do conteúdo pra
        // facilitar o diagnóstico e pula este e-mail, sem derrubar a execução.
        const inicio = bufferZip.slice(0, 200).toString('utf8').replace(/\s+/g, ' ').trim();
        console.error(`Aviso: o link de download do e-mail "${picked.subject}" não retornou um .zip válido (${bufferZip.length} bytes). Início do conteúdo: "${inicio}". Pulando este e-mail.`);
        continue;
      }
      const { nomeInterno, buffer } = extrairXlsxDoZip(bufferZip, 'relatório baixado do link do MOVIAS');
      bufferXlsx = buffer;
      nomeUsado = `link do MOVIAS → ${nomeInterno} (${bufferZip.length} bytes)`;
    } else if (picked.tipo === 'zip') {
      const { nomeInterno, buffer } = extrairXlsxDoZip(picked.attachment.content, picked.attachment.filename);
      bufferXlsx = buffer;
      nomeUsado = `${picked.attachment.filename} → ${nomeInterno}`;
    } else {
      bufferXlsx = picked.attachment.content;
      nomeUsado = `${picked.attachment.filename} (${picked.attachment.size} bytes)`;
    }
    console.log(`Processando e-mail "${picked.subject}" de ${picked.from} (${picked.date}). Anexo: ${nomeUsado}.`);

    try {
      const wb = XLSX.read(bufferXlsx, { type: 'buffer' });
      const { trips: tripsNovas, alertEvents: alertasNovos } = parseTelemetriaWorkbook(wb);
      tripsNovas.forEach(t => tripsMap.set(tripKey(t), t));
      alertasNovos.forEach(a => alertasMap.set(alertKey(a), a));
    } catch (err) {
      // Não derruba a execução inteira por causa de UM e-mail com formato
      // errado (ex.: o agendamento no MOVIAS às vezes manda o relatório de
      // "Alertas" em vez do relatório de "Eventos" que este robô espera —
      // problema de configuração do lado do MOVIAS, não deste código). Loga
      // como aviso e segue processando os demais e-mails encontrados.
      console.error(`Aviso: não consegui ler o anexo do e-mail "${picked.subject}" (${picked.date}) — pulando. Detalhe: ${err.message}`);
    }
  }

  const tripsFundidas = [...tripsMap.values()].sort((a, b) => new Date(a.DataDescarga) - new Date(b.DataDescarga));
  const alertasFundidos = [...alertasMap.values()].sort((a, b) => new Date(a.DataEvento) - new Date(b.DataEvento));

  const meta = buildMeta(tripsFundidas);
  const novoConteudo = { meta, trips: tripsFundidas, alertEvents: alertasFundidos };
  const anteriorComparavel = anterior ? { meta: anterior.meta, trips: anterior.trips, alertEvents: anterior.alertEvents } : null;
  const mudou = JSON.stringify(anteriorComparavel) !== JSON.stringify(novoConteudo);

  if (!mudou) {
    console.log('Nenhuma viagem nova em relação ao que já está publicado — nada a commitar.');
    return;
  }

  const payload = { ...novoConteudo, exportadoEm: new Date().toISOString() };
  fs.writeFileSync(OUTPUT_PATH, JSON.stringify(payload));
  const novas = tripsFundidas.length - tripsAnteriores.length;
  console.log(`dados_atuais.json atualizado: ${tripsFundidas.length} viagens no total (${novas >= 0 ? '+' + novas : novas} desde a última execução), ${meta.cacambas_monitoradas_piloto} caçambas (${meta.periodo_min}–${meta.periodo_max}).`);
}

main().catch(err => {
  console.error('Falha na automação:', err);
  process.exit(1);
});
