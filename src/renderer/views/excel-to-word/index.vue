<template>
  <section>
    <el-alert
      title="Excel转Word"
      description="讲Excel中的数据逐条转成Word模板的文件。"
      type="warning" :closable="false">
    </el-alert>
    <el-form :model="form" label-width="150px" class="mt-20">
      <el-form-item label="Excel数据表：">
        <el-upload
          v-if="form.excel == null"
          drag
          action="string"
          :show-file-list="false"
          :http-request="uploadExcel"
          accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel,.xlsx,.xls">
          <i class="el-icon-upload"></i>
          <div class="el-upload__text">将文件拖到此处，或<em>点击上传</em><br />只能上传Excel文件</div>
        </el-upload>
        <div v-else><i class="el-icon-success"></i></div>
      </el-form-item>
      <el-form-item label="Word模板：">
        <el-upload
          v-if="form.word == null"
          drag
          action="string"
          :show-file-list="false"
          :http-request="uploadWord"
          accept="application/msword,application/msword,.docx,.doc">
          <i class="el-icon-upload"></i>
          <div class="el-upload__text">将文件拖到此处，或<em>点击上传</em><br />只能上传Word文件</div>
        </el-upload>
        <div v-else><i class="el-icon-success"></i></div>
      </el-form-item>
      <el-form-item label="导出Word命名：" v-if="form.excel && form.word">
        <el-popover trigger="hover" content="规则：例*字段名**字段名*，即以*字段*包裹；如空值，则按系统默认规则。" placement="top">
          <el-input slot="reference" v-model="form.exportWordName" placeholder="请输入导出Word命名规则"></el-input>
        </el-popover>
      </el-form-item>
      <el-form-item class="text-align-right" v-if="form.excel && form.word">
        <el-progress :percentage="percentage" v-if="percentage != 0"></el-progress>
        <span ref="downloadURL"></span>
        <el-button @click="reset">重置</el-button>
        <el-button type="primary" @click="submit">立即转换</el-button>
      </el-form-item>
    </el-form>
  </section>
</template>

<script>
import createReport from 'docx-templates';
import JSZip from 'jszip';
import XLSX from 'xlsx';
import saveAs from './FileSaver.js';
import { trimAll } from '@/utils/pattern';

export default {
  name: 'ExcelToWord',
  data () {
    return {
      form: {
        excel: null,
        word: null,
        exportWordName: '',
      },
      percentage: 0,
    }
  },
  created(){
    console.log('home-created')
  },
  mounted(){
    console.log('home-mounted')
  },
  methods: {
    readFileIntoArrayBuffer(fd){//fd文件对象；讲温江对象转为buffer格式
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onerror = reject;
        reader.onload = () => {
          resolve(reader.result);
        };
        reader.readAsArrayBuffer(fd);
      });
    },
    async uploadExcel(item){
      //this.form.excel = await this.readFileIntoArrayBuffer(item.file);
      this.form.excel = item.file.path;//路径
    },
    async uploadWord(item){
      //this.form.word = await this.readFileIntoArrayBuffer(item.file);//有时报错待查
      this.form.word = item.file.path;//路径
    },
    reset(){
      this.form = {
        excel: null,
        word: null,
        exportWordName: '',
      };
      this.percentage = 0;
    },
    async submit(){
      /*const saveDataToFile = (data, fileName, mimeType) => {
        const blob = new Blob([data], { type: mimeType });
        const url = window.URL.createObjectURL(blob);
        downloadURL(url, fileName, mimeType);
        setTimeout(() => {
          window.URL.revokeObjectURL(url);
        }, 1000);
      };

      const downloadURL = (data, fileName) => {
        const a = document.createElement('a');
        a.href = data;
        a.download = fileName;
        document.body.appendChild(a);
        a.style = 'display: none';
        a.click();
        a.remove();
      };

      let report = await createReport({
          template: this.form.word,
          data: {'项目姓名':'asd','学号':1212},
        });

      saveDataToFile(
        report,
        'report.docx',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      );*/
      try{
        let workbook = XLSX.read(this.form.excel, {type: 'file'});
        let fromTo = '';
        let persons = [];
        for (let sheet in workbook.Sheets) {
          if (workbook.Sheets.hasOwnProperty(sheet)) {
            fromTo = workbook.Sheets[sheet]['!ref'];
            //console.log(fromTo);
            persons = persons.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]));
            break; // 如果只取第一张表，就取消注释这行
          }
        }
        //console.log(persons);

        var zip = new JSZip();//创建zip对象

        //excel逐条写入到word模板
        for(let i=0; i<persons.length; i++){
          let item = persons[i];
          for(let key in item){
            let val = item[key];
            let newKey = trimAll(key);
            delete item[key];
            item[newKey] = val;
          };
          //console.log(item,JSON.stringify(this.form.word))
          let report = await createReport({
            output: 'buffer',
            template: this.form.word,
            data: item,
          });
          if(this.form.exportWordName == ''){
            zip.file(`${i}.doc`, report);
          }
          else{
            let name = this.form.exportWordName.replace(/\*(.*?)\*/g,(word,matchWord,index)=>{
              let val = item[trimAll(matchWord)];
              val = val == null ? '' : val;
              return val;
            });
            //console.log(name)
            zip.file(`${name}.doc`, report);
          }
          this.percentage = Math.floor(i/(persons.length-1)*100);
        }

        setTimeout(()=>{
          //生成压缩包
          zip.generateAsync({type:"blob"}).then((content)=>{
            const url = window.URL.createObjectURL(content);
            this.$refs.downloadURL.innerHTML = `<a href="${url}" download="下载.zip" style="margin-right:20px;">下载</a>`;
            saveAs(content, "下载.zip");
          });
        },1000)
      }catch(err){
        alert('Error:'+err.message);
      }
    },
  },
  components: {
    
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped lang="scss">
/deep/ {
  .el-upload__text{
    line-height: 20px;
    font-size: 12px;
  }
}
</style>
