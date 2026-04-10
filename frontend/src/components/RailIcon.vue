<script setup lang="ts">
defineProps<{
  icon: string
}>()
</script>

<template>
  <svg class="app-rail__icon" viewBox="0 0 20 20" aria-hidden="true">
    <template v-if="icon === 'home'">
      <path d="M3.5 8.75 10 3.75l6.5 5V16a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1Z" />
      <path d="M7.75 17v-4.25h4.5V17" />
    </template>

    <template v-else-if="icon === 'bell'">
      <path d="M6 8.5a4 4 0 0 1 8 0V11l1.25 2.25H4.75L6 11Z" />
      <path d="M8.25 14.5a1.75 1.75 0 0 0 3.5 0" />
    </template>

    <template v-else-if="icon === 'search'">
      <circle cx="8.75" cy="8.75" r="4.75" />
      <path d="m12.5 12.5 3.25 3.25" />
    </template>

    <template v-else-if="icon === 'folder'">
      <path d="M3.5 6.5h4l1.5 1.75h7.5v6.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1Z" />
      <path d="M3.5 6.5v-.75a1 1 0 0 1 1-1h2.75L9 6.5" />
    </template>

    <template v-else-if="icon === 'grid'">
      <path d="M4 4.25h5v5H4Z" />
      <path d="M11 4.25h5v5h-5Z" />
      <path d="M4 10.75h5v5H4Z" />
      <path d="M11 10.75h5v5h-5Z" />
    </template>

    <template v-else-if="icon === 'clock'">
      <circle cx="10" cy="10" r="6.25" />
      <path d="M10 6.75v3.75l2.5 1.5" />
    </template>

    <template v-else-if="icon === 'star'">
      <path d="m10 3.75 1.75 3.75 4 .5-3 2.75.75 4-3.5-1.9-3.5 1.9.75-4-3-2.75 4-.5Z" />
    </template>

    <template v-else>
      <path d="M4 4.5h5v5H4ZM11 4.5h5v5h-5ZM4 11.5h5v5H4ZM11 11.5h5v5h-5Z" />
    </template>
  </svg>
</template>