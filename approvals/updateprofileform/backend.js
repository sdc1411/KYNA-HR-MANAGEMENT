const ConfigUpdateProfileForm = {
  ResponseDatabase: {
    WorkSheet: '1JHPMtOyR4EAttV97a0xbpKcutLrOJ_pnzIEpoQFZ44w',
    SheetName: 'Sheet1',
  },

}



class SubmitUpdateProfileForm {
  constructor() {
    this.db = SpreadsheetApp.openById(ConfigUpdateProfileForm.ResponseDatabase.WorkSheet)
    // this.pageSize = CONFIG.PAGE_SIZE
    // this.reverse = CONFIG.REVERSE
    // this.props = PropertiesService.getScriptProperties()
    // this.cache = CacheService.getScriptCache()
    this.headerId = "id"
  }

  createKeys(headers) {
    return headers.map(header => header.toString().trim())
  }

  createItemObject(keys, values) {
    const item = {}
    keys.forEach((key, index) => item[key] = values[index])
    return item
  }

  generateId(keys, records) {
    if (records.length === 0) return 1
    const indexOfId = keys.indexOf(this.headerId)
    if (indexOfId === -1) throw new Error(`"${this.headerId}" column is missing in the table!`)
    return records[records.length - 1][indexOfId] + 1
  }

  createValues(keys, item, values = []) {
    return keys.map((key, index) => {
      if (item.hasOwnProperty(key)) {
        return item[key]
      } else {
        return values[index] || null
      }
    })
  }

  submitUpdateProfileForm({ item }) {
    const ws = this.db.getSheetByName(ConfigUpdateProfileForm.ResponseDatabase.SheetName)
    const [headers, ...records] = ws.getDataRange().getValues()
    const keys = this.createKeys(headers)
    item.created_at = new Date()
    item.cb_status = "Waiting Approve"
    item.cmnd_Date = (item.cmndDate) ? convertDate(item.cmndDate) : null
    item.id = this.generateId(keys, records)
    const values = this.createValues(keys, item)
    ws.getRange(records.length + 2, 1, 1, values.length).setValues([values])
    return {
      success: true,
      message: `Yêu cầu cập nhật hồ sơ của bạn đã được gửi thành công đến Phòng Nhân Sự!`,
      data: item,
    }
  }

}

const appSubmitUpdateProfile = new SubmitUpdateProfileForm()

// lấy dữ liệu request update thông tin profile từ client
const createUpdateProfileForm = (params) => JSON.stringify(appSubmitUpdateProfile.submitUpdateProfileForm(JSON.parse(params)))



// const createUpdateProfileForm = (params) => {
//   const data = appSubmitUpdateProfile.submitUpdateProfileForm(JSON.parse(params))
//    console.log(params)
//    return JSON.stringify(data)
// }



