<script>
  import { get } from "svelte/store";
  import { loadData, data, failLoadData, failLoadMessage, openFile, FILE_PATH, getFileName } from "../stores/data";
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

    <p>Please ensure <code>{FILE_PATH}</code> is present and restart.</p>
  {:else}
    <h2>Menu</h2>

    {#if $data}
      <div class="content">
        <div>
          <p>Data succesfully imported from <code>{FILE_PATH}</code>.</p>
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
                  $changePage("review-data");
                }}
                disabled={$generatingEmails}
              >
                View Imported Data
              </button>
              <button
                on:click={() => {
                  openFile();
                }}
                disabled={$generatingEmails}
                class="btn btn-default"
                ><span class="fa fa-edit" /><span class="mrgn-lft-sm">Edit {getFileName()}</span>
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
                <span class="fa fa-save" /><span class="mrgn-lft-sm">Sync Data</span>
              </button>
            </div>
          </div>
        </div>

        <div class="mrgn-tp-lg">
          <h4>
            <!-- svelte-ignore a11y-invalid-attribute -->
            Email Generation
          </h4>
          <div class="flex mrgn-tp-md">
            <button
              class="btn btn-primary"
              on:click={() => {
                generateEmails(0);
              }}
              disabled={$generatingEmails}
            >
              Generate Employee Emails
            </button>
            <button
              class="btn btn-primary"
              on:click={() => {
                generateEmails(1);
              }}
              disabled={$generatingEmails}
            >
              Generate Manager Emails
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
