/**
 * https://github.com/google/clasp#readme
 * https://github.com/google/clasp/blob/master/docs/typescript.md
 * https://www.typescriptlang.org/docs/handbook/release-notes/typescript-2-0.html#non-null-assertion-operator
 * https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Nullish_coalescing
 * https://www.typescriptlang.org/tsconfig#strict
 * https://www.typescriptlang.org/tsconfig#alwaysStrict
 * https://typescript-eslint.io/getting-started
 */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function updateContactsSheetRun() {
  UpdateContactList.main();
}

const UpdateContactList = (function () {
  // https://medium.com/@Rahulx1/revealing-module-pattern-tips-e3442d4e352

  function updateContactsSheet(
    contactsSheet: GoogleAppsScript.Spreadsheet.Sheet,
    rowData: (string | undefined)[][],
    firstRow: number
  ) {
    const firstColumn = 1;
    const numRows = rowData.length;
    const numColumns = 7;

    contactsSheet.getRange(firstRow, firstColumn, numRows, numColumns)
      .setValues(rowData);

    return firstRow + numRows + 2;
  }

  function getContactsList(resourceName: string, resourceType: string) {
    // Advanced People Service
    // https://developers.google.com/apps-script/advanced/people

    // API
    // https://developers.google.com/people

    let ssData: (string | undefined)[][] = [];
    // https://support.google.com/googleapi/answer/7035610?hl=en
    const quotaUser = Session.getActiveUser().getEmail();
    const maxMembers = People.ContactGroups!.get(
      resourceName,
      { quotaUser }
    ).memberCount;
    const resourceNames = People.ContactGroups!.get(
      resourceName,
      { maxMembers, quotaUser }
    ).memberResourceNames;

    People.People!.getBatchGet(
      {
        resourceNames,
        personFields: [
          "names",
          "emailAddresses",
          "addresses",
          "phoneNumbers",
          "organizations"
        ],
        quotaUser
      }).responses
      ?.forEach(function (response) {
        const person = response.person;

        if (person) {
          // setup data to display in spreadsheet rows.
          ssData = [...ssData, [
            ((person.names)
              ? ((person.names[0]) ? person.names[0].displayNameLastFirst
                : "unk")
              : "unk"), // Name
            ((person.organizations)
              ? ((person.organizations[0])
                ? person.organizations[0].title
                : "")
              : ""), // Position
            resourceType, // Status
            ((person.phoneNumbers)
              ? ((person.phoneNumbers[0])
                ? person.phoneNumbers[0].value
                : "")
              : ""), // Phone
            ((person.addresses)
              ? ((person.addresses[0])
                ? [
                  person.addresses[0].streetAddress,
                  person.addresses[0].city,
                  person.addresses[0].region,
                  person.addresses[0].postalCode
                ].join(", ")
                : "")
              : ""), // Home Address
            ((person.emailAddresses)
              ? ((person.emailAddresses.length > 0)
                ? person.emailAddresses[0].value
                : "")
              : ""), // Primary Email
            ((person.emailAddresses)
              ? ((person.emailAddresses.length > 1)
                ? person.emailAddresses[1].value
                : "")
              : "") // Other Email
          ]];
        }
      });

    ssData.sort();

    return ssData;
  }

  function main() {
    const scriptProperties = PropertiesService.getScriptProperties();
    const contactsSpreadsheetID = scriptProperties.getProperty("CONTACTS_SPREADSHEET_ID");
    const activeSpreadsheet = ((contactsSpreadsheetID)
      ? SpreadsheetApp.openById(contactsSpreadsheetID)
      : SpreadsheetApp.getActive());
    const contactsListSheet = activeSpreadsheet.getSheetByName("Contact List")!;
    // https://developers.google.com/people/api/rest/v1/contactGroups/list
    const resourceNameObj = {
      active: scriptProperties?.getProperty("RESOURCE_NAME_ACTIVE") || "contactGroups/1cf9f5348e22c8b7",
      guest: scriptProperties?.getProperty("RESOURCE_NAME_GUEST") || "contactGroups/3c82995f899da957",
      inactive: scriptProperties?.getProperty("RESOURCE_NAME_INACTIVE") || "contactGroups/3a3fa8fc0d6be183",
      student: scriptProperties?.getProperty("RESOURCE_NAME_STUDENT") || "contactGroups/5d7c7a9d8e0c906d"
    };
    let ssData: (string | undefined)[][] = [];
    let rowCount = 2; // row 1 is the header
    const dt = new Date();

    // todo:  restore sheet if getContactsList fails?
    if (contactsListSheet.getLastRow() > 1) {
      contactsListSheet.getRange(2, 1, (contactsListSheet.getLastRow() - 1), 7)
        .clearContent();
    }

    ssData = getContactsList(resourceNameObj.active, "Active");
    rowCount = updateContactsSheet(contactsListSheet, ssData, rowCount);

    ssData = getContactsList(resourceNameObj.guest, "Guest");
    rowCount = updateContactsSheet(contactsListSheet, ssData, rowCount);

    ssData = getContactsList(resourceNameObj.student, "Student");
    rowCount = updateContactsSheet(contactsListSheet, ssData, rowCount);

    ssData = getContactsList(resourceNameObj.inactive, "Inactive");
    updateContactsSheet(contactsListSheet, ssData, rowCount);

    activeSpreadsheet.rename(`Sutherland Contacts ${dt.getFullYear()}`);
  }

  return { main };
}());
