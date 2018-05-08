<template>

  <div class="app-container" >
    <div style="margin:20px 0">
      <el-button 
        type="primary" 
        icon="el-icon-message" 
        @click="handleSendNotification()" 
        :loading="listLoading">
        開始寄出活動通知
      </el-button>
      <br>
      <el-progress style="left: 44%; transform: translateX(-50%);" type="circle" :width="200" :stroke-width="10" :percentage="progressSendNotification.progress" :status="progressSendNotification.status" v-show="loadingSendNotification"></el-progress>
    </div>
    
    <el-table :data="list" @selection-change="handleSelectionChange" :max-height=tableheight v-loading.body="listLoading ||loadingSendNotification " element-loading-text="正在為您拼命處理中" border fit highlight-current-row ref="multipleTable">
     

      <el-table-column
        type="index"
        width="80"
        align="center">
      </el-table-column>
       <el-table-column
        type="selection"
        width="70"        
        align="center">
      </el-table-column>
      <el-table-column
        label="活動通知"
        align="center"
        width="130">
          <template slot-scope="scope"> 
            <el-tag 
              :type=" scope.row.wasSentNotification ? 'success' : 'info' "               
            >
            {{ scope.row.wasSentNotification ? '已寄出' : '尚未' }}         
            </el-tag>
          </template>
      </el-table-column>
      <el-table-column
        label="已付款"
        align="center"
        width="130">
          <template slot-scope="scope"> 
            <el-tag 
              :type=" scope.row.wasPaid ? 'success' : 'info' " 
            >
            {{ scope.row.wasPaid ? '是' : '否' }}         
            </el-tag>
          </template>
      </el-table-column>
      
      <el-table-column label="訂單編號" align="center">
        <template slot-scope="scope">
          <span class="link-type" @click="handleView(scope.row)"> {{scope.row.guid}}</span>
        </template>
      </el-table-column>  

    
      <el-table-column label="聯絡人" align="center" width="200">
        <template slot-scope="scope">
          <span>{{scope.row.contact_name}}</span>
        </template>
      </el-table-column>
      <el-table-column label="聯絡信箱" width="250"  align="left"  >
        <template slot-scope="scope">
          <span>{{scope.row.contact_email}}</span>         
        </template>
      </el-table-column>
      
      <el-table-column
      label="操作"
      width="100"
       align="center">
      <template slot-scope="scope">
        <el-button
          @click.native.prevent="handleView(scope.row)"
          size="small"
          type="primary"
          >
          檢視
        </el-button>
        
          
      </template>
    </el-table-column>
    </el-table>

    <div class="pagination-container" v-show="false">
      <el-pagination background @size-change="handleSizeChange" @current-change="handleCurrentChange" :current-page="listQuery.page" :page-sizes="[3000]" :page-size="listQuery.limit" layout="total, sizes, prev, pager, next, jumper" :total="total">
      </el-pagination>
    </div>

    <!-- 編輯區塊 -->
    <el-dialog  title="詳細資訊"  :visible.sync="dialogFormVisible" width="70%" fullscreen >
      <el-form  ref="dataForm"  :model="tempData" label-position="left"  label-width="120px" style='color:#606266; width: 1000px; margin-left:50px;' v-loading="sendemailLoading || sendemailregisterLoading"
    element-loading-text="正在為您拼命處理"
    element-loading-spinner="el-icon-loading">
        <el-form-item label="訂單編號">        
          <span>{{tempData.guid}}</span>          
        </el-form-item>
        <el-form-item label="建立時間">        
          <span>{{tempData.CreatedTime}}</span>          
        </el-form-item>   
        <el-form-item label="公司資訊">          
          <el-collapse accordion>
            <el-collapse-item title="詳細資訊">
              <el-form-item label="公司抬頭">
                <el-input v-model="tempData.company_name" :disabled="!tempData.isEdit" size="medium"></el-input>
              </el-form-item>
              <br>
              <el-form-item label="統一編號">
                <el-input v-model="tempData.company_taxid" :disabled="!tempData.isEdit" size="medium"></el-input>
              </el-form-item>
              <br>
               <el-form-item label="發票形式">
                <el-input v-model="tempData.company_receipt" :disabled="!tempData.isEdit" size="medium"></el-input>
                
              </el-form-item>
            </el-collapse-item>
          </el-collapse>
        </el-form-item>
        <el-form-item label="聯絡人">          
          <el-collapse accordion>
            <el-collapse-item title="詳細資訊">
              <el-form-item label="姓名">
                <el-input v-model="tempData.contact_name" :disabled="!tempData.isEdit" size="medium"></el-input>
              </el-form-item>
              <br>
              <el-form-item label="聯絡電話">
                <el-input v-model="tempData.contact_phone" :disabled="!tempData.isEdit" size="medium"></el-input>
              </el-form-item>
              <br>
              <el-form-item label="分機">
                <el-input v-model="tempData.contact_ext" :disabled="!tempData.isEdit" size="medium"></el-input>
              </el-form-item>
              <br>
              <el-form-item label="電子郵件">
                <el-input v-model="tempData.contact_email" :disabled="!tempData.isEdit" size="medium"></el-input>
              </el-form-item>
              <br>
              <el-form-item label="縣市">
                <el-input v-model="tempData.contact_city" :disabled="!tempData.isEdit" size="medium"></el-input>
              </el-form-item>
              <br>
              <el-form-item label="行政區">
                <el-input v-model="tempData.contact_district" :disabled="!tempData.isEdit" size="medium"></el-input>
              </el-form-item>
              <br>
              <el-form-item label="完整街道地址">
                <el-input v-model="tempData.contact_address" :disabled="!tempData.isEdit" size="medium"></el-input>
              </el-form-item>
            </el-collapse-item>
          </el-collapse>
        </el-form-item>
        <el-form-item label="購買票種">          
          <el-collapse accordion>
            <el-collapse-item title="詳細資訊">
              <div v-for="(item, index) in tempData.Tickets" :key="index">                
                <el-form-item :label="detectTicketType(item.type)" >
                  <el-input-number v-model="item.count" :min="0" :disabled="!tempData.isEdit" ></el-input-number>
                </el-form-item>
                <br v-if="index != tempData.Tickets.length-1">
              </div>              
             
            </el-collapse-item>
          </el-collapse>
        </el-form-item>  
         
        <el-form-item label="備註">          
          <el-input type="textarea"
                    :autosize="{ minRows: 2, maxRows: 5}" 
                    v-model="tempData.remarks" 
                    :disabled="!tempData.isEdit" size="medium"
                    >
          </el-input>
        </el-form-item>   
        <el-form-item label="總票數">
          <el-input-number v-model="getTicketsCount" :min="0" :disabled="true" ></el-input-number>
        </el-form-item>   
        <el-form-item label="總金額">
          <el-input-number type="number" :min="0" v-model="tempData.tickets_amount" :disabled="!tempData.isEdit" controls-position="right"></el-input-number>
        </el-form-item>
        <el-form-item label="付款方式">          
           <el-radio-group v-model="tempData.payment_method" :disabled="!tempData.isEdit">
            <el-radio-button label="remit">匯款</el-radio-button>
            <el-radio-button label="credit">信用卡</el-radio-button>            
          </el-radio-group>
              <div class="remit" v-if="tempData.payment_method =='remit'">
                <hr>
                <el-form-item label="帳號後五碼">
                  <el-input v-model="tempData.payment_account" :disabled="!tempData.isEdit" size="medium" max="5"></el-input>
                </el-form-item>
              </div>
        </el-form-item>

     

        <el-form-item label="已付款">
          <el-tag :type="tempData.wasPaid ? 'success' : 'info'" size = "medium" >
            {{tempData.wasPaid? '是' : '否'}}
          </el-tag>
        </el-form-item>

        <el-form-item label="活動通知">
          <el-tag :type="tempData.wasSentNotification ? 'success' : 'info'" size = "medium" >
            {{tempData.wasSentNotification? '已寄出' : '尚未'}}
          </el-tag>
        </el-form-item>

        
      </el-form>
      <div slot="footer" class="dialog-footer">
             
        <el-button 
          @click="dialogFormVisible = false"
          v-show="!tempData.isEdit"
          :disabled="sendemailLoading||sendemailregisterLoading">
          關閉
          </el-button>        
      </div> 
    </el-dialog>
  </div>
</template>
<style scoped>
.demo-table-expand {
  font-size: 0;
}
.demo-table-expand label {
  width: 90px;
  color: #99a9bf;
}
.demo-table-expand .el-form-item {
  margin-right: 0;
  margin-bottom: 0;
  width: 30%;
}
.el-input.is-disabled .el-input__inner {
  background-color: #f5f7fa;
  border-color: #e4e7ed;
  color: #606266;
  cursor: not-allowed;
}
.el-input.is-disabled .el-input__inner,
.el-textarea.is-disabled .el-textarea__inner {
  background-color: #f5f7fa;
  border-color: #e4e7ed;
  color: #606266;
  cursor: not-allowed;
}
.el-radio-button__orig-radio:disabled + .el-radio-button__inner {
  color: #606266;
}

.el-input--medium {
  width: 280px;
}

.el-collapse-item__header {
  color: #888;
}
.el-dialog__footer {
  text-align: center;
}
.el-dialog.is-fullscreen {
  background-color: #fcfcfc;
}

/* hover */
.el-collapse-item__header:hover {
  color: #337ab7;
}

.el-collapse-item.is-active .el-collapse-item__header {
  color: #337ab7;
}
</style>
<script>
import {
  getList,
  putSendEmail,
  putSendEmail_Regitster,
  putSendEmail_Notification,
  putUpdate
} from '@/api/table'

export default {
  data() {
    return {
      list: null,
      listLoading: true,
      listQuery: {
        page: 1,
        limit: 3000,
        words: ''
      },
      total: 300,
      sendemailLoading: false,
      sendemailregisterLoading: false,
      tableheight: 0,
      dialogFormVisible: false,
      tempData: {
        Tickets: [
          {
            amount: 12000,
            count: 2,
            type: 0
          },
          {
            amount: 16000,
            count: 2,
            type: 1
          },
          {
            amount: 4000,
            count: 4,
            type: 2
          }
        ]
      },
      tempData_forUpdate: {},
      multipleSelection: [],
      progressSendNotification: {
        progress: 0,
        status: ''
      },
      loadingSendNotification: false
    }
  },
  filters: {
    statusFilter(status) {
      const statusMap = {
        published: 'success',
        draft: 'gray',
        deleted: 'danger'
      }
      return statusMap[status]
    }
  },
  created() {
    this.fetchData()
  },
  methods: {
    fetchData() {
      this.listLoading = true
      getList(this.listQuery).then(response => {
        this.list = response.data.filter(item => {
          return item.wasPaid === true
        })

        this.total = response.total
        this.listLoading = false
      })
    },
    getAppMainHeight() {
      this.tableheight = window.innerHeight - 200
    },
    // Toggle
    toggleSelection(rows) {
      console.log(1)
      if (rows) {
        rows.forEach(row => {
          this.$refs.multipleTable.toggleRowSelection(row)
        })
      } else {
        this.$refs.multipleTable.clearSelection()
      }
    },
    handleSelectionChange(val) {
      this.multipleSelection = val
    },

    handleSearch() {
      console.log(this.listQuery.words)
      this.fetchData()
    },
    handleView(row) {
      row.isEdit = false
      this.tempData = Object.assign({}, row) // copy obj
      this.dialogFormVisible = true
      // console.log('handleView', this.tempData)
    },
    handleStartEdit() {
      this.tempData_forUpdate = Object.assign({}, this.tempData)
      this.tempData.isEdit = true

      this.$message({
        showClose: true,
        type: 'warning',
        message: '開始編輯'
      })
      console.log('handleStartEdit', this.tempData)
    },
    handleCancelEdit() {
      this.$confirm('此操作將取消編輯, 資料將回到編輯前, 是否繼續?', '確認', {
        confirmButtonText: '確定',
        showCancelButton: false,
        type: 'warning'
      })
        .then(() => {
          this.tempData = Object.assign({}, this.tempData_forUpdate)
          this.tempData.isEdit = false
          this.$message({
            showClose: true,
            type: 'success',
            message: '已取消編輯'
          })
        })
        .catch(() => {
          this.$message({
            type: 'info',
            showClose: true,
            message: '已取消'
          })
        })
      console.log('handleCancelEdit', this.tempData)
    },
    handleUpdate(row) {
      this.$confirm('即將更改此筆資料, 是否繼續', '確認', {
        confirmButtonText: '確定',
        showCancelButton: false,
        type: 'warning'
      })
        .then(() => {
          this.sendemailLoading = true
          setTimeout(() => {
            putUpdate(this.tempData)
              .then(response => {
                console.log(response)
                this.sendemailLoading = false
                this.$message({
                  showClose: true,
                  type: 'success',
                  message: '修改成功'
                })

                getList(this.listQuery).then(response => {
                  this.list = response.data
                  this.total = response.total
                })

                this.tempData.isEdit = false
              })
              .catch(response => {
                console.log(response)
                this.sendemailLoading = false
                this.$message({
                  showClose: true,
                  duration: 5000,
                  type: 'error',
                  message: '資料錯誤, 請確認後送出'
                })
              })
            console.log(this.tempData)
          }, 1000)
        })
        .catch(() => {
          this.$message({
            type: 'info',
            message: '已取消'
          })
        })
    },
    detectTicketType(type) {
      var res = ''
      switch (type) {
        case 0:
          res = '早鳥票'
          break
        case 1:
          res = '一般票'
          break
        case 2:
          res = '學生票'
          break
      }
      return res
    },

    SendEmail(guid) {
      console.log(guid)

      this.$confirm('即將寄出付款成功的通知信件, 是否繼續', '確認', {
        confirmButtonText: '確定',
        showCancelButton: false,
        type: 'warning'
      })
        .then(() => {
          this.sendemailLoading = true
          putSendEmail(guid)
            .then(response => {
              console.log(response)
              this.sendemailLoading = false
              this.$message({
                showClose: true,
                type: 'success',
                message: '信件成功寄出'
              })

              getList(this.listQuery).then(response => {
                this.list = response.data
                this.total = response.total
              })
              this.tempData.wasSentLetter = true
            })
            .catch(() => {
              this.sendemailLoading = false
              this.$message({
                showClose: true,
                duration: 5000,
                type: 'error',
                message: '資料錯誤, 請聯絡工程師'
              })
            })
        })
        .catch(() => {
          this.$message({
            type: 'info',
            message: '已取消'
          })
        })
    },

    SendEmailRegister(guid) {
      this.$confirm('即將寄出報名成功的通知信件, 是否繼續', '確認', {
        confirmButtonText: '確定',
        showCancelButton: false,
        type: 'warning'
      })
        .then(() => {
          this.sendemailregisterLoading = true
          putSendEmail_Regitster(guid)
            .then(response => {
              console.log(response)
              this.sendemailregisterLoading = false
              this.$message({
                showClose: true,
                type: 'success',
                message: '信件成功寄出'
              })

              getList(this.listQuery).then(response => {
                this.list = response.data
                this.total = response.total
              })
              this.tempData.wasSentLetter = true
            })
            .catch(() => {
              this.sendemailregisterLoading = false
              this.$message({
                showClose: true,
                duration: 5000,
                type: 'error',
                message: '資料錯誤, 請聯絡工程師'
              })
            })
        })
        .catch(() => {
          this.$message({
            type: 'info',
            message: '已取消'
          })
        })
    },

    // pagination
    handleSizeChange(val) {
      this.listQuery.limit = val
      console.log('handleSizeChange', this.listQuery)
      this.fetchData()
      // this.getList()
    },

    handleCurrentChange(val) {
      this.listQuery.page = val
      console.log('handleCurrentChange', this.listQuery)
      this.fetchData()
    },
    waitFor(ms) {
      return new Promise((resolve, reject) => {
        setTimeout(() => {
          resolve()
        }, ms)
      })
    },
    async asyncForEach(array, callback) {
      for (let index = 0; index < array.length; index++) {
        await callback(array[index], index, array)
      }
    },

    async handleSendNotification() {
      this.progressSendNotification.progress = 0
      this.progressSendNotification.status = ''
      if (this.multipleSelection.length === 0) {
        this.$message({
          showClose: true,
          type: 'error',
          dangerouslyUseHTMLString: true,
          message: '<strong>未選取資料</strong>'
        })
        return
      }
      try {
        await this.$confirm('即將寄出活動的通知信件, 是否繼續', '確認', {
          confirmButtonText: '確定',
          showCancelButton: false,
          type: 'warning'
        })

        await this.asyncForEach(
          this.multipleSelection,
          // eslint-disable-next-line
          async (row, row_index) => {
            // await this.waitFor(2000)
            console.log(this.multipleSelection.length)
            this.loadingSendNotification = true
            await putSendEmail_Notification(row.guid)
              .then(response => {
                this.list.forEach(function(item, index) {
                  if (item.guid === row.guid) {
                    item.wasSentNotification = true
                  }
                })
                this.$message({
                  showClose: true,
                  type: 'success',
                  dangerouslyUseHTMLString: true,
                  message: '信件成功寄出<br>Mail to：' + row.contact_email
                })
                this.progressSendNotification.progress = Math.round(
                  (row_index + 1) / this.multipleSelection.length * 100
                )
              })
              .catch(() => {
                this.$message({
                  showClose: true,
                  duration: 5000,
                  type: 'error',
                  message: '信件寄出失敗'
                })
                this.loadingSendNotification = false
              })
            // console.log(row.guid)
          }
        )
        // this.sendemailregisterLoading = false
        console.log('Done')
        // this.progressSendNotification.progress = 100
        this.progressSendNotification.status = 'success'
        await this.waitFor(3000)
        this.loadingSendNotification = false
        this.$message({
          showClose: true,
          type: 'success',
          dangerouslyUseHTMLString: true,
          message: '全部處理完畢'
        })
      } catch (err) {
        console.log('err', err)
        if (err === 'cancel') {
          this.$message({
            type: 'info',
            showClose: true,
            message: '已取消'
          })
        } else {
          this.$message({
            type: 'error',
            showClose: true,
            message: '系統錯誤, 請聯繫工程師'
          })
        }
        this.loadingSendNotification = false
      }
    }
  },
  computed: {
    getTicketsCount: {
      get: function() {
        var res =
          this.tempData.Tickets[0].count +
          this.tempData.Tickets[1].count +
          this.tempData.Tickets[2].count
        return res
      },

      /* SET won't work */
      set: function() {
        var res =
          this.tempData.Tickets[0].count +
          this.tempData.Tickets[1].count +
          this.tempData.Tickets[2].count

        // console.log(res)
        return res
      }
    }
  },
  mounted() {
    this.getAppMainHeight()
  }
}
</script>
