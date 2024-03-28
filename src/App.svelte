<script lang="ts">
  import Loading from "./components/Loading.svelte";
  import { page, pageLoading } from "./stores/routes";

  window.resizeTo(screen.availWidth / 2, screen.availHeight);
  window.moveTo(-7, 0);
</script>

<main>
  <div class="container">
    <div class="mrgn-tp-lg">
      <h1 id="title">Mail Merge - Modified for Staffing Team</h1>
    </div>

    <svelte:component this={$page} />

    {#if $pageLoading}
      <Loading />
    {/if}
  </div>
</main>

<style>
  main {
    width: 100%;
    height: 100%;
    background-color: rgb(245, 245, 245);
  }

  #title {
    border-bottom: 3px solid #af3c43;
    margin-bottom: 0.2em;
    margin-top: 1em;
    padding-bottom: 0.2em;
  }
</style>
