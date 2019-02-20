<template>
  <section>

  </section>
</template>

<script>
export default {
  name: 'Home',
  data () {
    return {
      msg: 'This is home!',
    }
  },
  created(){
    console.log('home-created')
  },
  mounted(){
    console.log('home-mounted')
  },
  activated(){
    console.log('home-activated')
  },
  components: {
    
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped lang="scss">

</style>
