<template>
  Message:{{ message }}
  <input v-model="message"/>
  <button v-on:click="reverseMessage">反转</button>
</template>

<script>
export default {
  data() {
    return {
      message: "hello world!",
    };
  },
  methods: {
    reverseMessage() {
      this.message = this.message.split("").reverse().join("");
    },
  },
};
</script>

<style>
</style>