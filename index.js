import { GoogleSpreadsheet } from 'google-spreadsheet'
import fs from 'fs'
import path from 'path'
import excel from 'exceljs'

(async() => {
    const monthsMap = {
        'Janeiro': '01',
        'Fevereiro': '02',
        'Março': '03',
        'Abril': '04',
        'Maio': '05',
        'Junho': '06',
        'Julho': '07',
        'Agosto': '08',
        'Setembro': '09',
        'Outubro': '10',
        'Novembro': '11',
        'Dezembro': '12',
    }

    const credentialsBuffer = fs.readFileSync(path.resolve('credentials.json'))
    const credentials = JSON.parse(credentialsBuffer)

    const docIDs = {
        '2019': '19kIEKp77UpBQFRLo_slaBkjJa6HRltN6uDFPAVLEjIE',
        '2020': '19tZQaQl_W9LHD3KmkMD24C0Hsu2TpRQazruVy98oj5c',
        '2021': '11kX81CLna3N74HjEyUCU_GGlb1ffY_XhsCHMVCjMfgU' 
    }

    const doc = new GoogleSpreadsheet(docIDs['2021'])
    
    await doc.useServiceAccountAuth(credentials)
    await doc.loadInfo()

    const workbook = new excel.Workbook()
    const worksheet = workbook.addWorksheet(doc.title)

    worksheet.columns = [
        { header: 'Mês', key: 'month' },
        { header: 'Nome', key: 'name' },
        { header: 'Publicações', key: 'publications' },
        { header: 'Videos', key: 'videos' },
        { header: 'Horas', key: 'hours' },
        { header: 'Revisitas', key: 'returns' },
        { header: 'Estudos', key: 'studies' },
        { header: 'Observações', key: 'observations' },
    ]
    
    for (let count = 0; count < doc.sheetCount; count++) {
        const sheet = doc.sheetsByIndex[count]
        await sheet.setHeaderRow(['', '', 'name', 'publications', 'videos', 'hours', 'returns', 'studies', 'observations'])
        
        console.log('\n Sheet Name', sheet.a1SheetName)
        
        const [month, year] = sheet.a1SheetName.slice(1, sheet.a1SheetName.length -1).split('/')
        const monthYear = `${year}${monthsMap[month]}`

        try {
            const rows = await sheet.getRows({ offset: 3})
            rows.forEach(row => {
                if(row.name) {
                    worksheet.addRow({
                        month: monthYear, 
                        name: row.name,
                        publications: Number(row.publications),
                        videos: Number(row.videos | 0),
                        hours: Number(row.hours | 0),
                        returns: Number(row.returns | 0),
                        studies: Number(row.studies | 0),
                        observations: row.observations
                    });
                }
            })
        } catch(e) {
            console.log('Erro ao carregar linhas', e)
        }        
    }

    await workbook.xlsx.writeFile(`sheets/${doc.title} - ${Math.random()}.xlsx`)
})()