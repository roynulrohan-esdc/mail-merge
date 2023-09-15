<script>
  import { loadData, data, openFile, dataError, dataFilesList, readData, readDataFile, importData } from "../stores/data";
  import { config, configError, employeeEmails, generateEmails, generatingEmails, generationMessage, loadMailConfig, managerEmails, sendEmails, sendingEmails, sendingMessage } from "../stores/emails";
  import { changePage, pageLoading } from "../stores/routes";
  import { loadTemplates, templatesError, templatesList } from "../stores/templates";

  const getData = () => {
    $pageLoading = true;

    setTimeout(() => {
      loadTemplates();
      loadMailConfig();

      if ($config.sheet) {
        loadData().then(() => {
          $pageLoading = false;
        });
      }
    }, 200);
  };
  if (!$data) {
    getData();
  }

  let chosenTemplate,
    chosenDataFile,
    generationType = "Employee",
    sendType;

  $: {
    if (generationType) {
      chosenTemplate = "";
    }
  }
</script>

<div>
  {#if $dataError}
    <h2>Failed to load excel data</h2>

    <p>{$dataError.message}</p>

    <p>Path: <code>{$dataError.path}</code></p>
  {:else if $configError}
    <h2>Failed to load mail configuration</h2>

    <p><code>{$configError}</code></p>
  {:else if $templatesError}
    <h2>Failed to load templates</h2>

    <p>{$templatesError.message}</p>

    <p>Path: <code>{$templatesError.path}</code></p>
  {:else}
    <h2>Menu</h2>

    <div class="content">
      <div>
        <p>Sending from: <b>{$config.mailbox}</b></p>
      </div>

      <div class="mrgn-tp-lg">
        <h4>Import Data</h4>

        <div class="mrgn-tp-md">
          <label for="dataListDropdown">Choose an excel file to import from</label>
          <select id="dataListDropdown" class="form-control" bind:value={chosenDataFile} disabled={generationType === ""}>
            <option label={$dataFilesList.length === 0 ? "No data files found" : "Select a data file"} />
            {#each $dataFilesList as dataFile}
              <option value={dataFile}>{dataFile}</option>
            {/each}
          </select>
        </div>

        {#if chosenDataFile}
          <div class="flex mrgn-tp-lg">
            <button
              class="btn btn-primary"
              on:click={() => {
                employeeEmails.set([]);
                generationMessage.set("");
                sendingMessage.set("");
                importData(chosenDataFile);
              }}
              disabled={$generatingEmails}
            >
              Import Data
            </button>
          </div>
        {/if}
      </div>

      {#if $data}
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
                ><span class="fa fa-edit" /><span class="mrgn-lft-sm">Edit</span>
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
            <select id="typeDropdown" class="form-control" bind:value={generationType} disabled>
              <option label="Select a target" />
              <option value={"Employee"} selected>Employee</option>
              <option value={"Manager"}>Manager</option>
            </select>
          </div>

          <div class="mrgn-tp-md">
            <label for="templatesDropdown">Choose a template for script</label>
            <select id="templatesDropdown" class="form-control" bind:value={chosenTemplate} disabled={generationType === ""}>
              {#if generationType === "Employee"}
                <option label={$templatesList.employees.length === 0 ? "No employee templates found" : "Select a template"} />
                {#each $templatesList.employees as template}
                  <option value={template}>{template}</option>
                {/each}
              {:else if generationType === "Manager"}
                <option label={$templatesList.managers.length === 0 ? "No manager templates found" : "Select a template"} />
                {#each $templatesList.managers as template}
                  <option value={template}>{template}</option>
                {/each}
              {:else}
                <option label="Select a template" />
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

        {#if $employeeEmails.length !== 0}
          <div class="mrgn-tp-lg">
            <h4>Send Emails</h4>

            <div class="mrgn-tp-lg">
              <label for="typeDropdown">Employee or Manager</label>
              <select id="typeDropdown" class="form-control" bind:value={sendType} disabled>
                <option label="Select a target" />
                {#if $employeeEmails.length !== 0}
                  <option value={"Employee"} selected>Employee</option>
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
      {/if}
    </div>
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
