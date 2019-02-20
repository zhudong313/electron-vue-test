<template>
  <el-container class="is-vertical">
    <app-header></app-header>
    <el-container>
      <app-sidebar></app-sidebar>
      <app-main></app-main>
    </el-container>
  </el-container>
</template>

<script>
import { AppHeader, AppSidebar, AppMain } from './components'

export default {
  name: 'Layout',
  data () {
    return {
      
    }
  },
  components: {
    AppHeader, 
    AppSidebar, 
    AppMain,
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped lang="scss">

</style>
