<template>
  <el-aside width="200px">
    <el-menu default-active="0" class="sidebar" @open="handleOpen" @close="handleClose">

      <el-submenu v-for="(item,index) in navData" :index="index+''" :key="index"  v-if="item.children">

        <template slot="title">
          <span slot="title">{{ item.title }}</span>
        </template>
        <el-menu-item v-if="item.children" v-for="(child,childIndex) in item.children" :key="index+'-'+childIndex" :index="index+'-'+childIndex" @click="$router.push({ path: child.to })">
          {{ child.title }}
        </el-menu-item>
      </el-submenu>

      <el-menu-item class="siderbar-item" :index="index+''"  @click="$router.push({ path: item.to })" v-else>
        <i class="el-icon-menu"></i>
        <span slot="title">{{ item.title }}</span>
      </el-menu-item>

    </el-menu>

  </el-aside>
</template>

<script>
export default {
  name: 'Siderbar',
  data () {
    return {
      nav_pos:0,
      navData: [
        {
          title: 'Excel转Word',
          to: '/'
        },
        {
          title: '内容中心',
          to: '/a'
        },
        {
          title: '工单中心',
          to: '/b'
        },
        {
          title: '用户与权限管理',
          to: '/c'
        }
      ]
    }
  },
  methods:{
    handleOpen(){

    },
    handleClose(){

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
/deep/.sidebar{
  height: 100%;
  .el-submenu__title{
    height: auto;
  }
}
.siderbar-item{ padding-left: 15px !important;}
</style>
