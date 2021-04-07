<template>
<div class="primary">
  <v-content>
    <v-navigation-drawer
      permanent
      width="100%">
      <v-row
        class="fill-height"
        no-gutters>
        <v-col>
            <v-container class="main-content">
              <div class="header mt-5">
                Upload your file
              </div>
              <import-excel/>
            </v-container>
        </v-col>
      </v-row>
    </v-navigation-drawer>
  </v-content>
</div>
</template>

<script>
import ImportExcel from '@/components/ImportExcel'

export default {
  name: 'Home',
  components: {
    ImportExcel
  }
}
</script>
<style lang="scss" scoped>
.header {
    display: flex;
    justify-content: space-between;
    margin-bottom: 2rem;
    font-size: 28px;
    font-weight: bold;
}
.container {
  padding: 5rem 10rem;
}
.main-content {
  /* overflow-y: scroll; */
  height: 100vh;
}
.main-content::-webkit-scrollbar {
  width: 12px;
  background-color: #F5F5F5;
}

.main-content::-webkit-scrollbar-thumb {
  border-radius: 10px;
  background-color: #555;
}

.main-content::-webkit-scrollbar-track {
  border-radius: 10px;
  background-color: #F5F5F5;
}
.primary {
  height: 100%;
  .v-content {
      height: 100%;
  }
  .col {
      background-color: #f7f7f7;
  }
}
</style>
