// This is the function parseAllRoundsFromTable from dashboard.html
function parseAllRoundsFromTable(table, cat) {
  const result = {
    penyisihan: [],
    semifinal1: [],
    semifinal2: [],
    semifinal3: [],
    final: [],
    ranking_akhir: []
  };

  if (!table || !table.rows) return result;
  const rows = table.rows;

  // Cari baris-baris header tabel di tengah sheet secara dinamis
  const headerIndices = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (row && row.c && row.c[1] && row.c[1].v) {
      const valB = row.c[1].v.toString().trim().toUpperCase();
      if (valB === "KODE" || valB === "KODE REGU") {
        headerIndices.push(i);
      }
    }
  }

  console.log(`[${cat.toUpperCase()}] All matching header indices:`, headerIndices);

  // Saring header berurutan (ambil yang paling bawah jika berdampingan)
  const cleanHeaderIndices = [];
  for (let i = 0; i < headerIndices.length; i++) {
    const current = headerIndices[i];
    const next = headerIndices[i + 1];
    if (next === current + 1) {
      continue;
    }
    cleanHeaderIndices.push(current);
  }

  console.log(`[${cat.toUpperCase()}] Cleaned header indices:`, cleanHeaderIndices);

  const idxLolos = cleanHeaderIndices.length > 0 ? cleanHeaderIndices[0] : -1;
  const idxFinal = cleanHeaderIndices.length > 1 ? cleanHeaderIndices[1] : -1;

  console.log(`[${cat.toUpperCase()}] idxLolos (Semifinal header row):`, idxLolos);
  console.log(`[${cat.toUpperCase()}] idxFinal (Final header row):`, idxFinal);

  // A. BABAK PENYISIHAN (Baris 0 sampai sebelum idxLolos)
  const endPenyisihan = idxLolos !== -1 ? idxLolos : rows.length;
  for (let i = 0; i < endPenyisihan; i++) {
    const row = rows[i];
    if (row && row.c) {
      const kode = row.c[1] && row.c[1].v ? row.c[1].v.toString().trim() : '';
      const nama = row.c[2] && row.c[2].v ? row.c[2].v.toString().trim() : '';
      let totalRaw = 0;
      let rankRaw = 1;
      let status = '';
      
      if (cat === 'mula') {
        totalRaw = row.c[3] ? row.c[3].v : 0;
        rankRaw = row.c[4] ? row.c[4].v : 1;
        status = row.c[5] && row.c[5].v ? row.c[5].v.toString().trim() : '';
      } else { // 'madya' & 'wira'
        totalRaw = row.c[4] ? row.c[4].v : 0;
        rankRaw = row.c[5] ? row.c[5].v : 1;
        status = row.c[6] && row.c[6].v ? row.c[6].v.toString().trim() : '';
      }
      
      let undiRaw = '';
      let undiFinalRaw = '';
      if (cat === 'mula') {
        undiRaw = row.c[6] ? row.c[6].v : '';
        undiFinalRaw = row.c[7] ? row.c[7].v : '';
      } else if (cat === 'madya') {
        undiRaw = row.c[7] ? row.c[7].v : '';
        undiFinalRaw = row.c[8] ? row.c[8].v : '';
      } else if (cat === 'wira') {
        undiRaw = row.c[7] ? row.c[7].v : '';
        undiFinalRaw = row.c[8] ? row.c[8].v : '';
      }
      const undi = undiRaw !== null && undiRaw !== undefined ? undiRaw.toString().trim() : '';
      const undiFinal = undiFinalRaw !== null && undiFinalRaw !== undefined ? undiFinalRaw.toString().trim() : '';

      if (kode && kode !== '' && kode !== 'Kode') {
        let total = parseFloat(totalRaw);
        if (isNaN(total)) total = 0;
        let rank = parseInt(rankRaw);
        if (isNaN(rank)) rank = 1;

        result.penyisihan.push({
          kode: kode,
          nama: nama,
          total: total,
          rank: rank,
          status: status,
          undi: undi,
          undiFinal: undiFinal
        });
      }
    }
  }

  // B. BABAK SEMIFINAL I & II (Baris idxLolos + 1 sampai sebelum idxFinal)
  if (idxLolos !== -1) {
    const endSemifinal = idxFinal !== -1 ? idxFinal : rows.length;
    for (let i = idxLolos + 1; i < endSemifinal; i++) {
      const row = rows[i];
      if (row && row.c) {
        const kode = row.c[1] && row.c[1].v ? row.c[1].v.toString().trim() : '';
        const nama = row.c[2] && row.c[2].v ? row.c[2].v.toString().trim() : '';
        const grupRaw = row.c[3] ? row.c[3].v : 0; 
        const totalRaw = row.c[4] ? row.c[4].v : 0;
        const rankRaw = row.c[5] ? row.c[5].v : 1;
        const status = row.c[6] && row.c[6].v ? row.c[6].v.toString().trim() : '';

        if (kode && kode !== '' && kode !== 'Kode') {
          let total = parseFloat(totalRaw);
          if (isNaN(total)) total = 0;
          let rank = parseInt(rankRaw);
          if (isNaN(rank)) rank = 1;
          let grup = parseInt(grupRaw);
          
          const penyisihanInfo = result.penyisihan.find(p => p.kode === kode);
          const undiStr = penyisihanInfo ? penyisihanInfo.undi : '';
          const penyisihanRank = penyisihanInfo ? penyisihanInfo.rank : 99;

          if (isNaN(grup) || grup === 0) {
            if (undiStr !== '') {
              const uNum = parseInt(undiStr);
              grup = (uNum >= 1 && uNum <= 5) ? 1 : 2;
            } else {
              const lolosList = result.penyisihan
                .filter(p => p.status === 'LOLOS SEMIFINAL')
                .sort((a, b) => a.rank - b.rank);
              const idxLolos = lolosList.findIndex(p => p.kode === kode);
              if (idxLolos !== -1) {
                grup = (idxLolos < 5) ? 1 : 2;
              } else {
                grup = 1;
              }
            }
          }

          const item = {
            kode: kode,
            nama: nama,
            total: total,
            rank: rank,
            status: status,
            undi: undiStr,
            penyisihanRank: penyisihanRank
          };

          if (grup === 1) {
            result.semifinal1.push(item);
          } else if (grup === 2) {
            result.semifinal2.push(item);
          } else if (grup === 3) {
            result.semifinal3.push(item);
          }
        }
      }
    }
  }

  // C. BABAK FINAL & RANKING AKHIR (Baris idxFinal + 1 sampai akhir)
  if (idxFinal !== -1) {
    for (let i = idxFinal + 1; i < rows.length; i++) {
      const row = rows[i];
      if (row && row.c) {
        const kode = row.c[1] && row.c[1].v ? row.c[1].v.toString().trim() : '';
        const nama = row.c[2] && row.c[2].v ? row.c[2].v.toString().trim() : '';
        const totalRaw = row.c[3] ? row.c[3].v : 0;
        const rankRaw = row.c[4] ? row.c[4].v : 1;
        const status = row.c[5] && row.c[5].v ? row.c[5].v.toString().trim() : '';

        if (kode && kode !== '' && kode !== 'Kode') {
          let total = parseFloat(totalRaw);
          if (isNaN(total)) total = 0;
          let rank = parseInt(rankRaw);
          if (isNaN(rank)) rank = 1;

          const penyisihanInfo = result.penyisihan.find(p => p.kode === kode);
          const undiFinalStr = penyisihanInfo ? penyisihanInfo.undiFinal : '';
          const penyisihanRank = penyisihanInfo ? penyisihanInfo.rank : 99;

          const item = {
            kode: kode,
            nama: nama,
            total: total,
            rank: rank,
            status: status || 'FINALIS',
            undiFinal: undiFinalStr,
            penyisihanRank: penyisihanRank
          };

          result.final.push(item);
          result.ranking_akhir.push({...item});
        }
      }
    }
  }

  return result;
}

const spreadsheetId = "1V0TX1XM8sT2nto1S0AhfQEDNC7g4D6cTLg1sOWxiZ0U"; // MAIN SPREADSHEET

async function runTest() {
  for (const cat of ['mula', 'madya', 'wira']) {
    const sheetName = cat === 'mula' ? 'PMR MULA' : (cat === 'madya' ? 'PMR MADYA' : 'PMR WIRA');
    console.log(`\n-----------------------------------------`);
    console.log(`FETCHING DATA FOR CATEGORY: ${cat.toUpperCase()}`);
    console.log(`-----------------------------------------`);
    try {
      const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}&t=${Date.now()}`;
      const res = await fetch(url);
      const text = await res.text();
      
      // GViz returns a callback or prefix we need to strip
      const match = text.match(/google\.visualization\.Query\.setResponse\((.*)\);/s);
      if (!match) {
        console.error("Failed to parse GViz JSON response wrapper.");
        continue;
      }
      
      const payload = JSON.parse(match[1]);
      const table = payload.table;
      
      const result = parseAllRoundsFromTable(table, cat);
      console.log(`Penyisihan count:`, result.penyisihan.length);
      console.log(`Semifinal 1 count:`, result.semifinal1.length);
      console.log(`Semifinal 2 count:`, result.semifinal2.length);
      console.log(`Semifinal 3 (Tie Break) count:`, result.semifinal3.length);
      console.log(`Final count:`, result.final.length);
      if (result.final.length > 0) {
        console.log(`Finalists:`, result.final.map(f => `${f.kode}: ${f.nama} (${f.total} pts, rank ${f.rank})`));
      }
    } catch (e) {
      console.error(`Error processing category ${cat}:`, e);
    }
  }
}

runTest();
