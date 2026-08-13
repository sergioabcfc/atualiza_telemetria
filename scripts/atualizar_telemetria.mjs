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

async function encontrarAnexoMaisRecente(client) {
  const lock = await client.getMailboxLock('INBOX');
  try {
    const since = new Date();
    since.setDate(since.getDate() - LOOKBACK_DAYS);

    const uids = await client.search({ since }, { uid: true });
    if (!uids || uids.length === 0) {
      console.log(`Nenhum e-mail encontrado nos últimos ${LOOKBACK_DAYS} dia(s).`);
      return null;
    }

    const candidatos = [];
    for await (const msg of client.fetch(uids, { envelope: true, uid: true }, { uid: true })) {
      const fromAddr = (msg.envelope && msg.envelope.from && msg.envelope.from[0] && msg.envelope.from[0].address || '').toLowerCase();
      if (SENDER_FILTER && !fromAddr.includes(SENDER_FILTER)) continue;
      candidatos.push({ uid: msg.uid, date: msg.envelope?.date, subject: msg.envelope?.subject, from: fromAddr });
    }
    candidatos.sort((a, b) => new Date(b.date) - new Date(a.date));

    for (const cand of candidatos) {
      const { content } = await client.download(cand.uid, undefined, { uid: true });
      const parsed = await simpleParser(content);
      const anexoXlsx = (parsed.attachments || []).find(a =>
        /\.xlsx?$/i.test(a.filename || '') ||
        a.contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      if (anexoXlsx) {
        return { ...cand, attachment: anexoXlsx };
      }
    }
    console.log(`Nenhum dos ${candidatos.length} e-mail(s) encontrados trazia um anexo .xlsx.`);
    return null;
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

  let picked;
  try {
    picked = await encontrarAnexoMaisRecente(client);
  } finally {
    await client.logout();
  }

  if (!picked) {
    console.log('Nada a atualizar nesta execução.');
    return;
  }

  console.log(`Usando e-mail "${picked.subject}" de ${picked.from} (${picked.date}). Anexo: ${picked.attachment.filename} (${picked.attachment.size} bytes).`);

  const wb = XLSX.read(picked.attachment.content, { type: 'buffer' });
  const { trips: tripsNovas, alertEvents: alertasNovos } = parseTelemetriaWorkbook(wb);

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
  tripsNovas.forEach(t => tripsMap.set(tripKey(t), t));
  const tripsFundidas = [...tripsMap.values()].sort((a, b) => new Date(a.DataDescarga) - new Date(b.DataDescarga));

  const alertasMap = new Map(alertasAnteriores.map(a => [alertKey(a), a]));
  alertasNovos.forEach(a => alertasMap.set(alertKey(a), a));
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
