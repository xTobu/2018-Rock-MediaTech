<template>

  <div class="app-container" >
    
    <div style="margin:20px 0">      
    <!-- <el-button @click="toggleSelection([list[1], list[2]])">切换第二、第三行的选中状态</el-button>
    <el-button @click="toggleSelection()">取消选择</el-button> -->
      <a :href="`/NPOI/DownloadExcel`">    
        <el-button>下載總表單</el-button>
      </a>
      <el-input
        placeholder="请输入内容"
        prefix-icon="el-icon-search"
        v-model="listQuery.words"
        style="width:330px"        
        @keyup.enter.native="handleSearch()"

      >
       <el-button slot="append" icon="el-icon-search" @click="handleSearch()" ></el-button>
      </el-input>
    
  </div>
    <el-table :data="list" @selection-change="handleSelectionChange" :max-height=tableheight v-loading.body="listLoading" element-loading-text="Loading" border fit highlight-current-row ref="multipleTable">
      <!-- <el-table-column type="expand">
        <template slot-scope="scope">
          <el-form label-position="left" inline class="demo-table-expand">
            <el-form-item >
              <h3>Tickets</h3>
            </el-form-item>
            <div v-for="(item, index) in scope.row.Tickets" :key="index">                
                <el-form-item label="票種" >
                  <span> {{item.type}}</span>
                </el-form-item>
                
                <el-form-item label="數量" >
                  <span> {{item.count}}</span>
                </el-form-item> 
            </div>
          </el-form>
        </template>
      </el-table-column> -->
     <!-- <el-table-column
      type="index"
      width="50"
      align="center">
    </el-table-column> -->

      <el-table-column label="公司抬頭" align="center" width="250">
        <template slot-scope="scope">
          <span>{{scope.row.company_name}}</span>
        </template>
      </el-table-column>
    
      <el-table-column label="聯絡人" align="center">
        <template slot-scope="scope">
          <span>{{scope.row.contact_name}}</span>
        </template>
      </el-table-column>

      <el-table-column label="聯絡信箱" width="280"  align="left"  >
        <template slot-scope="scope">
          <span class="link-type" @click="handleView(scope.row)"> {{scope.row.contact_email}}</span>         
        </template>
      </el-table-column>
      <el-table-column label="聯絡電話" width="180" align="center">
        <template slot-scope="scope">
          <span>{{scope.row.contact_phone}}{{scope.row.contact_ext?" ext. "+scope.row.contact_ext:""}} </span>
        </template>
      </el-table-column>
      <el-table-column label="付款方式" width="80" align="center">
        <template slot-scope="scope">
          <span>{{scope.row.payment_method=="remit"?"匯款":"信用卡"}}</span>
        </template>
      </el-table-column>
      <el-table-column label="帳號後五碼" width="130" align="center">
        <template slot-scope="scope">
          <span>{{scope.row.payment_account}}</span>
        </template>
      </el-table-column>
      <el-table-column label="總金額" width="80" align="center">
        <template slot-scope="scope">
          <span>{{scope.row.tickets_amount}}</span>
        </template>
      </el-table-column>
   
      <!-- <el-table-column class-name="status-col" label="Status" width="110" align="center">
        <template slot-scope="scope">
          <el-tag :type="scope.row.status | statusFilter">{{scope.row.status}}</el-tag>
        </template>
      </el-table-column> -->
      <!-- <el-table-column align="center" prop="created_at" label="建立時間" width="200"  >
        <template slot-scope="scope">         
          <span>{{scope.row.CreatedTime}}</span>
        </template>
      </el-table-column> -->
      <el-table-column
      label="已付款通知"
      align="center"
      width="100">
        <template slot-scope="scope"> 
          <el-tag 
            :type=" scope.row.wasSentLetter ? 'success' : 'info' " 
          >
          {{ scope.row.wasSentLetter ? '已寄出' : '尚未' }}         
          </el-tag>
        </template>
    </el-table-column>

      <el-table-column
      label="操作"
      width="130"
       align="center">
      <template slot-scope="scope">
        <el-button
          @click.native.prevent="handleView(scope.row)"
          size="small"
          type="primary"
          >
          檢視
        </el-button>
        <!-- <el-button
          v-show=false
          @click.native.prevent="deleteRow(scope.$index, list)"
          size="mini"
          >
          移除
        </el-button> -->
        <!-- <el-button
          v-show=false
          @click.native.prevent="open_check_sendemail(scope.row.guid)"
          size="mini">
          寄信
        </el-button> -->
        <!-- <el-tooltip v-show=false :content="'Switch value: ' + scope.row.wasPaid" placement="top">
          <el-switch
            v-model="scope.row.wasPaid"
            :disabled = scope.row.wasPaid
            :change="SwitchOnChange()"
            active-color="#13ce66"
            inactive-color="#ff4949"
          >
          </el-switch>
        </el-tooltip> -->
          
      </template>
    </el-table-column>
    </el-table>

    <div class="pagination-container">
      <el-pagination background @size-change="handleSizeChange" @current-change="handleCurrentChange" :current-page="listQuery.page" :page-sizes="[10,20,50]" :page-size="listQuery.limit" layout="total, sizes, prev, pager, next, jumper" :total="total">
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
                <el-input v-model="detectReceipt" :disabled="!tempData.isEdit" size="medium"></el-input>
                
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

        <el-form-item label="報名成功通知">
          <el-tag :type="tempData.wasSentRegister ? 'success' : 'info'" size = "medium" >
            {{tempData.wasSentRegister? '已寄出' : '尚未'}}
          </el-tag>
        </el-form-item>

        <el-form-item label="已付款通知">
          <el-tag :type="tempData.wasSentLetter ? 'success' : 'info'" size = "medium" >
            {{tempData.wasSentLetter? '已寄出' : '尚未'}}
          </el-tag>
        </el-form-item>

        
      </el-form>
      <div slot="footer" class="dialog-footer">
        <el-button 
          :type="tempData.isEdit ? 'warning' : 'primary'" 
          @click="tempData.isEdit ? handleUpdate() : handleStartEdit()"
          :disabled="sendemailLoading||sendemailregisterLoading">
          {{tempData.isEdit ? '更新' : '編輯'}}
        </el-button>
        <el-button 
          type="info" 
          v-show="tempData.isEdit"
          @click="handleCancelEdit()">
         取消編輯
        </el-button>

      <el-button 
        v-if="!tempData.isEdit"
        type="success"  
        :loading="sendemailregisterLoading" 
        @click="SendEmailRegister(tempData.guid)"
        :disabled="sendemailLoading||sendemailregisterLoading"
        >
        {{tempData.wasSentRegister ? "重新寄出報名成功通知" : "寄出報名成功通知"}}
      </el-button>

      <el-button 
        v-if="!tempData.isEdit"
        type="success"  
        :loading="sendemailLoading" 
        @click="SendEmail(tempData.guid)"
        :disabled="sendemailLoading||sendemailregisterLoading"
        >
        {{tempData.wasSentLetter ? "重新寄出已付款通知" : "寄出已付款通知"}}
      </el-button>
      
        
        <el-button 
          @click="dialogFormVisible = false"
          v-show="!tempData.isEdit"
          :disabled="sendemailLoading||sendemailregisterLoading"
        >
          關閉</el-button>        
      </div> 
    </el-dialog>
  </div>
</template>
<style>
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
  putUpdate
} from '@/api/table'

export default {
  data() {
    return {
      list: null,
      listLoading: true,
      listQuery: {
        page: 1,
        limit: 10,
        words: ''
      },
      total: 300,
      sendemailLoading: false,
      sendemailregisterLoading: false,
      tableheight: 0,
      dialogFormVisible: false,
      tempData: {
        company_receipt: 'double',
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
      tempData_forUpdate: {}
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
        console.log(response)
        this.list = response.data
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
      console.log(this.multipleSelection)
    },
    deleteRow(index, rows) {
      alert('delete')
      // rows.splice(index, 1)
    },

    SwitchOnChange() {},
    handleSearch() {
      console.log(this.listQuery.words)
      this.fetchData()
    },
    handleView(row) {
      row.isEdit = false
      this.tempData = Object.assign({}, row) // copy obj
      this.dialogFormVisible = true
      console.log('handleView', this.tempData)
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
            console.log(this.tempData)
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
      // this.getList()
    }
    // getTicketsCount() {
    //   var res =
    //     this.tempData.Tickets[0].count +
    //     this.tempData.Tickets[1].count +
    //     this.tempData.Tickets[2].count

    //   console.log(res)
    //   return res
    //   // loop over array with parameter, reference in template generateSomething('myParam')
    // }
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

        console.log(res)
        return res
      }
    },
    detectReceipt: {
      get: function() {
        var res
        console.log(this.tempData)
        switch (this.tempData.company_receipt) {
          case 'double':
            res = '二聯式'
            break
          case 'triple':
            res = '三聯式'
            break
          default:
            res = ''
            break
        }

        return res
      },
      set: function() {
        var res
        switch (this.tempData.company_receipt) {
          case 'double':
            res = '二聯式'
            break
          case 'triple':
            res = '三聯式'
            break
          default:
            res = ''
            break
        }
        console.log(res)
        return res
      }
    }
  },
  mounted() {
    this.getAppMainHeight()
  }
}
</script>
