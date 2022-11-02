<script>
  import { loadData, data, failLoadData, failLoadMessage } from "../stores/data";
  import { changePage, pageLoading } from "../stores/routes";

  $pageLoading = true;

  setTimeout(() => {
    if (!$data) loadData();

    $pageLoading = false;
  }, 200);
</script>

<div>
  {#if $failLoadData}
    <h2>Failed to load data from excel</h2>

    <p><code>{$failLoadMessage}</code></p>

    <p>Please place <code>Scenario1.xlsx</code>, <code>Scenario2.xlsx</code> in the <code>/input/</code> directory and restart.</p>
  {:else}
    <h2>Menu</h2>

    {#if $data}
      <div class="content">
        <div>
          <p>Data succesfully imported from <code>Scenario1.xlsx</code> & <code>Scenario2.xlsx</code>.</p>
        </div>

        <div class="mrgn-tp-lg">
          <h4>
            <!-- svelte-ignore a11y-invalid-attribute -->
            1. Review Data before proceeding
          </h4>

          <div class="mrgn-tp-md">
            <button
              class="btn btn-default"
              on:click={() => {
                $changePage("review-data");
              }}
            >
              Review
            </button>
          </div>
        </div>

        <div class="mrgn-tp-lg">
          <h4>
            <!-- svelte-ignore a11y-invalid-attribute -->
            2. Generate Emails
          </h4>
          <div class="flex mrgn-tp-md">
            <button
              class="btn btn-default"
              on:click={() => {
                // TODO
              }}
            >
              Contact Employees
            </button>
            <button
              class="btn btn-default"
              on:click={() => {
                // TODO
              }}
            >
              Contact Managers
            </button>
          </div>
        </div>
      </div>
    {/if}
  {/if}
</div>

<style>
  h2 {
    color: grey;
    width: 100%;
    text-align: center;
  }

  .content {
    margin-top: 50px;
  }

  .flex {
    display: flex;
  }

  .flex > * {
    margin-right: 20px;
  }
</style>
