<script>
  import Loading from "./components/Loading.svelte";
  import { page, pageLoading } from "./stores/routes";

  window.resizeTo(screen.availWidth / 2, screen.availHeight);
  window.moveTo(-7, 0);
</script>

<main>
  <div class="container">
    <div class="mrgn-tp-lg">
      <h1 id="title">Mail Merge</h1>
    </div>
  </div>

  <svelte:component this={$page} />

  {#if $pageLoading}
    <Loading />
  {/if}
</main>

<style>
  main {
    width: 100%;
    display: flex;
    flex-direction: column;
    align-items: center;
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
