<script>
  import { get } from "svelte/store";
  import { loadData, data, failLoadData, failLoadMessage, openFile } from "../stores/data";
  import { generateEmails, generatingEmails, generationMessage } from "../stores/emails";
  import { changePage, pageLoading } from "../stores/routes";

  const getData = () => {
    $pageLoading = true;

    setTimeout(() => {
      loadData().then(() => {
        $pageLoading = false;
      });
    }, 200);
  };

  if (!$data) {
    getData();
  }
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
            Change Data
          </h4>

          <div class="mrgn-tp-md">
            <div class="flex mrgn-tp-md">
              <button
                class="btn btn-default"
                on:click={() => {
                  openFile(1);
                }}
                disabled={$generatingEmails}
              >
                Open Scenario1.xlsx
              </button>
              <button
                class="btn btn-default"
                on:click={() => {
                  openFile(2);
                }}
                disabled={$generatingEmails}
              >
                Open Scenario2.xlsx
              </button>
            </div>
          </div>

          <div class="mrgn-tp-md">
            <p>If data has been changed, sync to update data</p>
            <div class="flex mrgn-tp-md">
              <button
                class="btn btn-default"
                on:click={() => {
                  getData();
                }}
                disabled={$generatingEmails}
              >
                Sync Data
              </button>
              <button
                class="btn btn-default"
                on:click={() => {
                  $changePage("review-data");
                }}
                disabled={$generatingEmails}
              >
                View Imported Data
              </button>
            </div>
          </div>
        </div>

        <div class="mrgn-tp-lg">
          <h4>
            <!-- svelte-ignore a11y-invalid-attribute -->
            Generate Emails
          </h4>
          <div class="flex mrgn-tp-md">
            <button
              class="btn btn-default"
              on:click={() => {
                generateEmails(0);
              }}
              disabled={$generatingEmails}
            >
              Contact Employees
            </button>
            <button
              class="btn btn-default"
              on:click={() => {
                generateEmails(1);
              }}
              disabled={$generatingEmails}
            >
              Contact Managers
            </button>
          </div>

          <div class="mrgn-tp-lg">
            {#if $generatingEmails}
              <p>Generating emails...</p>
            {/if}

            {#if !$generatingEmails && $generationMessage}
              <p>{$generationMessage.message} <code>{$generationMessage.path}</code></p>
            {/if}
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
