let Excel = require('exceljs');
let fs = require('fs');

(async () => {
  let workbook = new Excel.Workbook()
  let filename = `${__dirname}/batch2.xlsx`
  await workbook.xlsx.readFile(filename);
  let worksheet = workbook.getWorksheet("Sheet1")

  for (let i = 2; i <= worksheet.rowCount; i++) {
    let nim = worksheet.getCell('A' + i).value
    let url = worksheet.getCell('B' + i).value

    if (typeof url == 'object') url = url.text

    worksheet.getCell('C' + i).value = i;
    worksheet.getRow(i).commit();

    if (isValidHttpUrl(url)) {
      crawl(url, async function (e) {
        // worksheet.getCell('C' + i).value = e;
        // console.log(nim, worksheet.getCell('C' + i).value)
        fs.appendFile('result.csv', `${nim},${e}\n`, function (err) {
          if (err) throw err;
        });
      }, nim)
    }
  }
})()

const Crawler = require("crawler")

// console.log(worksheet[0].data)
// let i = 0;
// worksheet[0].data.map(async e => {
//   // console.log(e)
//   let link = e[1]
//   if (!isValidHttpUrl(link)) return
//   //
//   crawl(link, worksheet, i++)
// })

function crawl(url, f, nim) {
  let skor;

  let c = new Crawler({
    rateLimit: 1000,
    callback: function (error, res, done) {
      let fakultas = [
        'ftik.unikom.ac.id',
        'feb.unikom.ac.id',
        'fisip.unikom.ac.id',
        'fh.unikom.ac.id',
        'fd.unikom.ac.id',
        'fib.unikom.ac.id',
        'fs.unikom.ac.id',
      ];

      let prodi = [
        'msi.pasca.unikom.ac.id',
        'mds.pasca.unikom.ac.id',
        'dkv.unikom.ac.id',
        'di.unikom.ac.id',
        'mm.pasca.unikom.ac.id',
        'if.unikom.ac.id',
        'is.unikom.ac.id',
        'mi.unikom.ac.id',
        'sk.unikom.ac.id',
        'tk.unikom.ac.id',
        'ti.unikom.ac.id',
        'ar.unikom.ac.id',
        'pwk.unikom.ac.id',
        'elektro.unikom.ac.id',
        'sipil.unikom.ac.id',
        'mn.unikom.ac.id',
        'mp.unikom.ac.id',
        'kp.unikom.ac.id',
        'ak.unikom.ac.id',
        'ka.unikom.ac.id',
        'ik.unikom.ac.id',
        'ip.unikom.ac.id',
        'hi.unikom.ac.id',
        'dg.unikom.ac.id',
        'si.unikom.ac.id',
        'sj.unikom.ac.id',
        'hk.unikom.ac.id',
      ];

      let utama = [
        'unikom.ac.id'
      ];

      let basePoint = 15;
      let readPoint = 10; // Max 60
      let linkPoint = 5; // Max 15
      skor = basePoint

      if (error) {
        console.log(error);
      } else {
        let $ = res.$;

        let match = false
        let readTime = null
        let re = null
        let alphabet = 'abcdefghijklmnopqrstuvwxyz'.split('');
        let listClass = [];

        alphabet.map(y => {
          alphabet.map(z => {
            listClass.push(`${y}${z}`)
          })
        })

        for(let c = 0; c < listClass.length; c++) {
          let _class = listClass[c];
          readTime = $('.' + _class).parent().html();
          if (!readTime) continue
          re = /(?<time>\d+) min read/;
          match = readTime.match(re)
          if (match) break
        }

        if (!match) {
          console.log(nim, 'Cant Find Min Read')
          return
        }

        let actualReadTime = Number(readTime?.match(re)[1])
        if (actualReadTime >= 8) skor += 5
        if (actualReadTime > 6) actualReadTime = 6
        skor += actualReadTime * readPoint;

        let links = $("article").find('a[href]');
        let skorLink = 0;
        links.map(e => {
          let href = links[e]?.attribs?.href;

          href = href.replace('https://', '')
          href = href.replace('http://', '')
          href = href.replace(/^[\\/]+|[\\/]+$/g, '')

          let cekFakultas = fakultas.includes(href)
          if (cekFakultas) skorLink += linkPoint

          let cekProdi = prodi.includes(href)
          if (cekProdi) skorLink += linkPoint

          let cekUtama = utama.includes(href)
          if (cekUtama) skorLink += linkPoint
        })
        if (skorLink > 15) skorLink = 15;
        skor += skorLink

        console.log(skor)
        f(skor)
      }
      done();
    }
  })
  c.queue(url);
}

function isValidHttpUrl(string) {
  let url;

  try {
    url = new URL(string);
  } catch (_) {
    return false;
  }

  return url.protocol === "http:" || url.protocol === "https:";
}