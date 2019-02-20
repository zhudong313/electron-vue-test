<template>
  <el-header class="header">
    <el-row class="h">
      <el-col class="h d-flex" :span="6">
        <h1 class="logo d-flex align-items-center">ZDD - 软件</h1>
      </el-col>
      <el-col :span="18"></el-col>
    </el-row>
  </el-header>
</template>

<script>
export default {
  name: 'AppHeader',
  data () {
    return {

    }
  },
  components: {
    
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped lang="scss">
@import '@/styles/variable.scss';
@import '@/styles/mixin.scss';
.header{
  background: $bg_header;
  border-bottom: #e6e6e6 1px solid;
  .logo{
    color: $color_blue;
  }
}

</style>
