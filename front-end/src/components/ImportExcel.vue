<template>
  <div>
    <h3>Import XLSX</h3>
    <input type="file" @change="onChange" />
    <button type="button" @click="exportExcel()">Export</button>
  </div>
</template>

<script>
import XLSX from 'xlsx'
export default {
  name: 'ImportExcel',

  data () {
    return {
      headerFile: ['Person ID', 'Name', 'Department', 'Date', 'Check-in', 'TemperatureIn', 'Check-out', 'TemperatureOut', 'Abnormal'],
      dateKey: [],
      data: []
    }
  },

  methods: {
    onChange (e) {
      const file = e.target.files[0]
      if (file) {
        this.upload(file)
      }
    },

    async upload (file) {
      const reader = new FileReader()
      reader.onload = (e) => {
        const arrayBuff = e.target.result
        /* if binary string, read with type 'binary' */
        const wb = XLSX.read(arrayBuff, { type: 'array' })
        const wsname = wb.SheetNames[0]
        const ws = wb.Sheets[wsname]
        /* Convert array of arrays */
        const dataRead = XLSX.utils.sheet_to_json(ws, { header: 1 })
        for (let i = 1; i < dataRead.length; i++) {
          this.groupData(dataRead[i])
        }
      }
      reader.readAsArrayBuffer(file)
    },

    groupData (dataRead) {
      const row = this.convertDateTime(dataRead)
      const date = `${row[3].d}/${row[3].m}/${row[3].y}`
      const time = `${row[3].H}:${row[3].M}:${row[3].S}`
      const empKey = row[1]

      if (!this.data[empKey]) {
        this.data[empKey] = {}
      }
      if (!this.dateKey[date]) {
        this.dateKey[date] = date
      }

      if (!this.data[empKey][date]) {
        this.data[empKey][date] = {
          'Person ID': row[0],
          Name: row[1],
          Department: row[2],
          Date: date,
          'Check-in': time,
          TemperatureIn: row[9],
          'Check-out': null,
          TemperatureOut: null,
          Abnormal: row[10]
        }
      } else {
        this.data[empKey][date]['Check-out'] = time
        this.data[empKey][date].TemperatureOut = row[9]
      }
    },

    convertDateTime (data) {
      data[3] = XLSX.SSF.parse_date_code(data[3]) // 3 is index of time in file excel
      return data
    },

    formatData (data) {
      const dataFormat = []
      const keyData = Object.keys(data)
      const date = Object.keys(this.dateKey)
      for (let i = 0; i < keyData.length; i++) {
        for (let j = 0; j < date.length; j++) {
          const dataObj = data[keyData[i]][date[j]]
          if (dataObj) {
            dataFormat.push(dataObj)
          }
        }
      }

      return dataFormat
    },

    exportExcel () {
      const data = this.formatData(this.data)
      /* convert state to workbook */
      const ws = XLSX.utils.json_to_sheet(data, { header: this.headerFile })
      const wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, ws, 'SheetJS')
      // /* generate file and send to client */
      XLSX.writeFile(wb, 'report.xlsx')
    }
  }
}
</script>

<style>
</style>
