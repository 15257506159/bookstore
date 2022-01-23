<template>
    Clock:{{ hour }}:{{ minute }}:{{ second }}
</template>

<script>
export default {
  name: "ClockTest",
  data() {
    return {
      time: 0,
      hour:0,
      minute:0,
      second:0,
    };
  },
  mounted() {
    setInterval(() => {
      this.time++;
      this.hour=parseInt(this.time/3600%24)
      this.minute=parseInt(this.time/60%60)
      this.second=this.time%60
    }, 1000);
  },
};
</script>

<style>
</style>