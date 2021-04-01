<template>
  <div class="home">
    <import-excel/>
  </div>
</template>

<script>
// @ is an alias to /src
// import HelloWorld from '@/components/HelloWorld.vue'
import ImportExcel from '@/components/ImportExcel'

export default {
  name: 'Home',
  components: {
    ImportExcel
  }
}
</script>
