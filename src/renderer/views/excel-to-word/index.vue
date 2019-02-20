<template>
  <section>
    <el-alert
      title="Excel转Word"
      description="讲Excel中的数据逐条转成Word模板的文件。"
      type="warning" :closable="false">
    </el-alert>
    <el-form :model="form" label-width="100px" class="mt-20">
      <el-form-item label="Excel数据表：">
        <el-upload
          v-if="form.excel == null"
          drag
          action="string"
          :show-file-list="false"
          :http-request="uploadExcel"
          accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel,.xlsx">
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
          accept="application/msword,application/msword,.docx">
          <i class="el-icon-upload"></i>
          <div class="el-upload__text">将文件拖到此处，或<em>点击上传</em><br />只能上传Word文件</div>
        </el-upload>
        <div v-else><i class="el-icon-success"></i></div>
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

export default {
  name: 'ExcelToWord',
  data () {
    return {
      form: {
        excel: null,
        word: null,
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
      this.form.excel = await this.readFileIntoArrayBuffer(item.file);
    },
    async uploadWord(item){
      this.form.word = await this.readFileIntoArrayBuffer(item.file);
    },
    reset(){
      this.form = {
        excel: null,
        word: null,
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
      };*/

      /*saveDataToFile(
        report,
        'report.docx',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      );*/

      let workbook = XLSX.read(this.form.excel, {type: 'buffer'});
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
        let report = await createReport({
          template: this.form.word,
          data: item,
        });
        zip.file(`${i}.doc`, report);
        this.percentage = Math.floor(i/(persons.length-1)*100);
      }
      //生成压缩包
      zip.generateAsync({type:"blob"}).then((content)=>{
        const url = window.URL.createObjectURL(content);
        this.$refs.downloadURL.innerHTML = `<a href="${url}" download="下载.zip">下载</a>`;
        saveAs(content, "example.zip");
      });
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
