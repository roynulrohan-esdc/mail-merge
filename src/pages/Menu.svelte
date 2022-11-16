<script>
  import { loadData, data, failLoadData, failLoadMessage, openFile, FILE_PATH, getFileName } from "../stores/data";
  import { employeeEmails, generateEmails, generatingEmails, generationMessage, managerEmails, sendEmails, sendingEmails, sendingMessage } from "../stores/emails";
  import { changePage, pageLoading } from "../stores/routes";
  import { loadTemplates, templatesError, templatesList } from "../stores/templates";

  const getData = () => {
    $pageLoading = true;

    setTimeout(() => {
      loadTemplates();
      loadData().then(() => {
        $pageLoading = false;
      });
    }, 200);
  };
  if (!$data) {
    getData();
  }

  let chosenTemplate, generationType, sendType;

  $: {
    if (generationType) {
      chosenTemplate = "";
    }
  }
</script>

<div>
  {#if $failLoadData}
    <h2>Failed to load data from excel</h2>

    <p><code>{$failLoadMessage}</code></p>

    <p>Please ensure <code>{FILE_PATH}</code> is present and restart.</p>
  {:else if $templatesError}
    <h2>Failed to load templates</h2>

    <p>{$templatesError.message}</p>

    <p>Path: <code>{$templatesError.path}</code></p>
  {:else}
    <h2>Menu</h2>

    {#if $data}
      <div class="content">
        <div>
          <p>Data succesfully imported from <code>{FILE_PATH}</code>.</p>
        </div>

        <div class="mrgn-tp-lg">
          <h4>Change Data</h4>

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
          <h4>Email Generation</h4>

          <div class="mrgn-tp-lg">
            <label for="typeDropdown">Employee or Manager</label>
            <select id="typeDropdown" class="form-control" bind:value={generationType}>
              <option label="Select a target" />
              <option value={"Employee"}>Employee</option>
              <option value={"Manager"}>Manager</option>
            </select>
          </div>

          <div class="mrgn-tp-md">
            <label for="templatesDropdown">Choose a template for script</label>
            <select id="templatesDropdown" class="form-control" bind:value={chosenTemplate} disabled={generationType === ""}>
              <option label="Select a template" />
              {#if generationType === "Employee"}
                {#each $templatesList.employees as template}
                  <option value={template}>{template}</option>
                {/each}
              {:else if generationType === "Manager"}
                {#each $templatesList.managers as template}
                  <option value={template}>{template}</option>
                {/each}
              {/if}
            </select>
          </div>

          {#if chosenTemplate && generationType}
            <div class="flex mrgn-tp-lg">
              <button
                class="btn btn-primary"
                on:click={() => {
                  generateEmails(generationType === "Employee" ? 0 : 1, chosenTemplate);
                }}
                disabled={$generatingEmails}
              >
                Generate Emails
              </button>
            </div>
          {/if}

          <div class="mrgn-tp-lg">
            {#if $generatingEmails}
              <p>Generating emails...</p>
            {/if}

            {#if !$generatingEmails && $generationMessage}
              <p>
                {$generationMessage.message}
                {#if $generationMessage.path}
                  <code>{$generationMessage.path}</code>
                {/if}
              </p>
            {/if}
          </div>
        </div>

        {#if $managerEmails.length !== 0 || $employeeEmails.length !== 0}
          <div class="mrgn-tp-lg">
            <h4>Send Emails</h4>

            <div class="mrgn-tp-lg">
              <label for="typeDropdown">Employee or Manager</label>
              <select id="typeDropdown" class="form-control" bind:value={sendType}>
                <option label="Select a target" />
                {#if $employeeEmails.length !== 0}
                  <option value={"Employee"}>Employee</option>
                {/if}
                {#if $managerEmails.length !== 0}
                  <option value={"Manager"}>Manager</option>
                {/if}
              </select>
            </div>

            {#if sendType}
              <div class="flex mrgn-tp-lg">
                <button
                  class="btn btn-primary"
                  on:click={() => {
                    sendEmails(sendType === "Employee" ? 0 : 1);
                  }}
                  disabled={$generatingEmails}
                >
                  Send Emails
                </button>
              </div>
            {/if}

            <div class="mrgn-tp-lg">
              {#if $sendingEmails}
                <p>Sending emails...</p>
              {/if}

              {#if !$sendingEmails && $sendingMessage}
                <p>
                  {$sendingMessage.message}
                  {#if $sendingMessage.path}
                    <code>{$sendingMessage.path}</code>
                  {/if}
                </p>
              {/if}
            </div>
          </div>
        {/if}
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
