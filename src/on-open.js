//function to display a menu bar where funtions can be executed
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Command')
      .addSubMenu(ui.createMenu('For Create')
        .addItem('Create Schedule and send an email', 'create_schedule')
      )
      .addSubMenu(ui.createMenu('For Update/ Delete')
        .addItem('Get target events', 'show_search_type_input_modal')
        .addSeparator()
        .addItem('Update/ Delete designated events','update_delete_events')
      )
      .addToUi();
}


