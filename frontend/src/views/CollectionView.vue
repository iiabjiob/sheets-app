<script setup lang="ts">
import { computed } from 'vue'
import { useRoute } from 'vue-router'

const route = useRoute()
const pageTitle = computed(() => String(route.meta.pageTitle ?? 'Page'))
const pageDescription = computed(() => String(route.meta.pageDescription ?? ''))
const pageKind = computed(() => String(route.meta.pageKind ?? 'generic'))
</script>

<template>
  <section class="collection-view" :data-kind="pageKind">
    <div class="collection-view__hero">
      <p class="collection-view__eyebrow">Library</p>
      <h1>{{ pageTitle }}</h1>
      <p class="collection-view__description">{{ pageDescription }}</p>
    </div>

    <div class="collection-view__panel">
      <h2>{{ pageTitle }} is being split into a dedicated surface</h2>
      <p>
        Workspaces now own the main editing shell. This section stays available for
        cross-workspace exploration, shortcuts, and future browsing flows.
      </p>
    </div>
  </section>
</template>

<style scoped>
.collection-view {
  min-height: 100%;
  display: grid;
  align-content: start;
  gap: 20px;
  padding: 32px;
}

.collection-view__hero {
  display: grid;
  gap: 8px;
  max-width: 720px;
}

.collection-view__eyebrow {
  margin: 0;
  font-size: 12px;
  font-weight: 700;
  letter-spacing: 0.12em;
  text-transform: uppercase;
  color: var(--color-text-muted);
}

.collection-view__hero h1,
.collection-view__panel h2 {
  margin: 0;
}

.collection-view__description,
.collection-view__panel p {
  margin: 0;
  max-width: 64ch;
  color: var(--color-text-muted);
  line-height: 1.6;
}

.collection-view__panel {
  display: grid;
  gap: 10px;
  max-width: 760px;
  padding: 22px 24px;
  border: 1px solid var(--color-border);
  border-radius: var(--radius-lg);
  background:
    linear-gradient(180deg, color-mix(in srgb, var(--color-surface) 92%, white) 0%, var(--color-surface) 100%);
  box-shadow: 0 18px 40px rgba(15, 23, 42, 0.06);
}
</style>