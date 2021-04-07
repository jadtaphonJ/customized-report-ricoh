<template>
  <div>
    <v-card class="content-container">
      <div class="import-container">
      <v-row>
        <div class="col-12">
          File should be .CSV
        </div>
        <div class="col-6">
          <input type="file" @change="onChange" />
        </div>
        <div class="col-6">
          <v-btn
            @click="exportExcel()"
            outlined
            class="mr-2"
            width="180"
            color="primary">
            Export
          </v-btn>
        </div>
      </v-row>
      </div>
    </v-card>
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
      data: [],
      wscols: [
        { wch: 15 }, // Person ID
        { wch: 20 }, // Name
        { wch: 20 }, // Department
        { wch: 15 }, // Date
        { wch: 10 }, // Check-in
        { wch: 15 }, // TemperatureIn
        { wch: 10 }, // Check-out
        { wch: 15 }, // TemperatureOut
        { wch: 10 } // Abnormal
      ]
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

    convertTimeToStr (hour, minute) {
      if (hour < 10) {
        hour = `0${hour}`
      }
      if (minute < 10) {
        minute = `0${minute}`
      }

      return `${hour}:${minute}`
    },

    groupData (dataRead) {
      const row = this.convertDateTime(dataRead)
      const date = `${row[3].d}/${row[3].m}/${row[3].y}`
      const time = this.convertTimeToStr(row[3].H, row[3].M)
      const empKey = row[1]

      if (!this.data[empKey]) {
        this.data[empKey] = {}
      }
      if (!this.dateKey[date]) {
        this.dateKey[date] = date
      }

      if (!this.data[empKey][date]) {
        this.data[empKey][date] = {
          'Person ID': row[0].slice(1),
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

    async exportExcel () {
      // const requestOptions = {
      //   method: 'GET',
      //   headers: { 'Content-Type': 'application/json' }
      // }
      // body: JSON.stringify({ title: "Vue POST Request Example" })
      // const data = this.formatData(this.data)
      const response = await fetch('http://localhost:3548/get-report')
      const data = await response.json()
      console.log(data)
      /* convert state to workbook */
      // const ws = XLSX.utils.json_to_sheet(data, { header: this.headerFile })
      // ws['!cols'] = this.wscols
      // const wb = XLSX.utils.book_new()
      // XLSX.utils.book_append_sheet(wb, ws, 'SheetJS')
      // // /* generate file and send to client */
      // XLSX.writeFile(wb, 'report.xlsx', { bookSST: true })
    }
  }
}
</script>

<style lang="scss" scoped>
.import-container {
    padding: 3.7rem;
}
.content-container {
  box-shadow: 0 5px 25px 0 rgba(0, 0, 0, 0.05);
  border-radius: 15px !important;
  .v-data-table__wrapper {
    border-radius: 15px 15px 0px 0px !important;
    thead.v-data-table-header {
      background-color: #6c6c6c;
      span, i {
          font-weight: normal;
          font-size: 15px !important;
          color: white !important;
      }
      i {
          padding: 5px;
      }
    }
  }
}
</style>
