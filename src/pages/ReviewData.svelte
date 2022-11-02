<script>
  import { loadData, data, failLoadData, failLoadMessage } from "../stores/data";
  import { changePage, pageLoading } from "../stores/routes";

  let selected = "";
</script>

<div class="main">
  <div class="return">
    <button
      on:click={() => {
        $changePage("menu");
      }}
    >
      Return
    </button>
  </div>

  <h2>Review Excel Data</h2>

  <div class="content">
    <div>
      <label for="fileDropdown">Choose a file to review</label>
      <select id="fileDropdown" bind:value={selected}>
        <option label="File Select" />
        <option value={"Scenario1.xlsx"}>Scenario1.xlsx</option>
        <option value={"Scenario2.xlsx"}>Scenario2.xlsx</option>
      </select>
    </div>
    {#if $data}
      <div class="mrgn-tp-lg">
        {#if selected === "Scenario1.xlsx"}
          <table class="table results-table table-bordered table-striped">
            <thead>
              <tr>
                <th scope="col">First Name</th>
                <th scope="col">Last Name</th>
                <th scope="col">Email</th>
              </tr>
            </thead>
            <tbody>
              {#each $data.scenarioOne as row}
                <tr>
                  <td>
                    {row.firstName}
                  </td>
                  <td>
                    {row.lastName}
                  </td>
                  <td>
                    {row.email}
                  </td>
                </tr>
              {/each}
            </tbody>
          </table>
        {:else if selected === "Scenario2.xlsx"}
          <table class="table results-table table-bordered table-striped">
            <thead>
              <tr>
                <th scope="col">Full Name</th>
                <th scope="col">Email</th>
                <th scope="col">Supervisor's Name</th>
                <th scope="col">Manager Email</th>
              </tr>
            </thead>
            <tbody>
              {#each $data.scenarioTwo as row}
                <tr>
                  <td>
                    {row.fullName}
                  </td>
                  <td>
                    {row.email}
                  </td>
                  <td>
                    {row.supervisorName}
                  </td>
                  <td>
                    {row.supervisorEmail}
                  </td>
                </tr>
              {/each}
            </tbody>
          </table>
        {/if}
      </div>
    {/if}
  </div>
</div>

<style>
  .main {
    position: relative;
  }

  .return {
    position: absolute;
  }

  .return button {
    padding: 5px 14px;
    background-color: #eaebed;
    border: 1px solid #dcdee1;
    border-radius: 4px;
    color: rgb(0, 52, 82);
  }

  .return button:hover {
    background-color: #cccccc;
  }

  h2 {
    color: grey;
    width: 100%;
    text-align: center;
  }

  .content {
    margin-top: 50px;
  }

  label {
    text-align: center;
    font-size: large;
  }

  select {
    width: fit-content;
  }

  .results-table {
    width: 100%;
    font-size: 14px;
  }

  .results-table * {
    white-space: nowrap;
  }

  .results-table th,
  .results-table td {
    text-align: center;
  }
</style>
