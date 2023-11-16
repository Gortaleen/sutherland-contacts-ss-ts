/**
 * https://github.com/Gortaleen/sutherland-contacts-ss-ts
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
  UpdateContactList.main("forceUpdate");
}

interface ContactGroupResponse
  extends GoogleAppsScript.People.Schema.ContactGroupResponse {}
interface PersonResponse
  extends GoogleAppsScript.People.Schema.PersonResponse {}

const UpdateContactList = (function () {
  // https://medium.com/@Rahulx1/revealing-module-pattern-tips-e3442d4e352

  /**
   * Compares the time the band contacts sheet was last updated to the last
   * updated times of Google Contact groups (labelled as: Active, Guest, ...)
   * and also the times that indivdual contacts were updated (name, phone, ...).
   */
  function personnelChanges(
    ssLastUpdated: GoogleAppsScript.Base.Date,
    scriptProperties: GoogleAppsScript.Properties.Properties,
    ...peopleResponses: Array<Array<PersonResponse> | undefined>
  ): boolean {
    // first check to see if any contacts have been added or deleted
    const connectionsSyncToken = scriptProperties.getProperty(
      "CONNECTIONS_SYNC_TOKEN"
    );
    // https://developers.google.com/people/api/rest/v1/people.connections/list
    const listConnectionsResponse = People.People?.Connections?.list(
      "people/me",
      {
        personFields: ["names", "metadata"],
        requestSyncToken: true,
        syncToken: connectionsSyncToken,
      }
    );
    const totalPeople = listConnectionsResponse?.totalPeople || 0;

    if (listConnectionsResponse?.nextSyncToken) {
      scriptProperties.setProperty(
        "CONNECTIONS_SYNC_TOKEN",
        listConnectionsResponse.nextSyncToken || ""
      );
    }

    // contact was either added or deleted
    if (totalPeople > 0) {
      return true;
    }

    // next check for edits to contacts
    return peopleResponses.some(function (peopleResponse) {
      return peopleResponse?.some(function (personResponse) {
        let personLastUpdatedStr;
        let personLastUpdated;
        if (personResponse.person?.metadata?.sources) {
          personLastUpdatedStr =
            personResponse.person.metadata.sources[0].updateTime || "";
          personLastUpdated = new Date(personLastUpdatedStr);
          return personLastUpdated > ssLastUpdated;
        }
        return false;
      });
    });
  }

  /**
   * Gets contacts' data (e.g. name, phone, ...) from Google Contacts app.
   */
  function getPeopleResponses(
    contactGroupResponse: Array<ContactGroupResponse> | undefined,
    quotaUser: string
  ): Array<Array<PersonResponse> | undefined> {
    let actives: Array<PersonResponse> | undefined;
    let guests: Array<PersonResponse> | undefined;
    let inactives: Array<PersonResponse> | undefined;
    let students: Array<PersonResponse> | undefined;

    contactGroupResponse?.forEach(function (groupResponse) {
      const resourceNames = groupResponse.contactGroup?.memberResourceNames;
      const groupName = groupResponse.contactGroup?.name?.toUpperCase();
      const personFields = [
        "addresses",
        "emailAddresses",
        "metadata",
        "names",
        "organizations",
        "phoneNumbers",
      ];
      let response;

      if (groupName) {
        // https://developers.google.com/people/api/rest/v1/people/getBatchGet
        response = People.People?.getBatchGet({
          resourceNames,
          personFields,
          quotaUser,
        });

        if (response) {
          switch (groupName) {
            case "ACTIVE":
              actives = response.responses;
              break;
            case "GUEST":
              guests = response.responses;
              break;
            case "INACTIVE":
              inactives = response.responses;
              break;
            case "STUDENT":
              students = response.responses;
              break;
            default:
              break;
          }
        }
      }
    });

    return [actives, guests, inactives, students];
  }

  /**
   * Gets Google Contacts organized by labels (e.g., Active, Guest, ...)
   */
  function getContactGroups(
    quotaUser: string,
    ...resourceNames: Array<string | null>
  ): Array<ContactGroupResponse> | undefined {
    const maxMembers = 1000; // ? this value is arbitrary
    let responseBody;

    if (resourceNames.length > 0) {
      // https://developers.google.com/people/api/rest/v1/contactGroups/batchGet
      responseBody = People.ContactGroups?.batchGet({
        resourceNames,
        maxMembers,
        quotaUser,
      });
    }

    return responseBody?.responses;
  }

  /**
   * Function currying a particular Google Sheet with the setValues function
   */
  function addToSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    return function (
      rowData: Array<Array<string | undefined>>,
      firstRow: number
    ) {
      const firstColumn = 1;
      const numRows = rowData.length;
      const numColumns = 7;

      if (rowData.length > 0) {
        sheet
          .getRange(firstRow, firstColumn, numRows, numColumns)
          .setValues(rowData);

        return firstRow + numRows + 2;
      }

      return firstRow;
    };
  }

  /**
   * Formats contacts' data (e.g., name, phone, ...) into rows to be filed to a
   * Google spreadsheet.
   */
  function getContactsData(
    resourceType: string,
    personResponses: Array<PersonResponse> | undefined
  ) {
    let ssData: Array<Array<string | undefined>> = [];

    personResponses?.forEach(function (response) {
      const person = response.person;

      if (person) {
        // setup data to display in spreadsheet rows.
        ssData = [
          ...ssData,
          [
            person.names
              ? person.names[0]
                ? person.names[0].displayNameLastFirst
                : "unk"
              : "unk", // Name
            person.organizations
              ? person.organizations[0]
                ? person.organizations[0].title
                : ""
              : "", // Position
            resourceType, // Status
            person.phoneNumbers
              ? person.phoneNumbers[0]
                ? person.phoneNumbers[0].value
                : ""
              : "", // Phone
            person.addresses
              ? person.addresses[0]
                ? [
                    person.addresses[0].streetAddress,
                    person.addresses[0].city,
                    person.addresses[0].region,
                    person.addresses[0].postalCode,
                  ].join(", ")
                : ""
              : "", // Home Address
            person.emailAddresses
              ? person.emailAddresses.length > 0
                ? person.emailAddresses[0].value
                : ""
              : "", // Primary Email
            person.emailAddresses
              ? person.emailAddresses.length > 1
                ? person.emailAddresses[1].value
                : ""
              : "", // Other Email
          ],
        ];
      }
    });
    ssData.sort();

    return ssData;
  }

  function main(forceUpdate = "") {
    // https://support.google.com/googleapi/answer/7035610?hl=en
    const quotaUser = Session.getActiveUser().getEmail();
    const scriptProperties = PropertiesService.getScriptProperties();
    const contactsSpreadsheetID = scriptProperties.getProperty(
      "CONTACTS_SPREADSHEET_ID"
    );
    const activeSpreadsheet = contactsSpreadsheetID
      ? SpreadsheetApp.openById(contactsSpreadsheetID)
      : SpreadsheetApp.getActive();
    const ssLastUpdated = DriveApp.getFileById(
      activeSpreadsheet.getId()
    ).getLastUpdated();
    const contactsListSheet = activeSpreadsheet.getSheetByName("Contact List")!;
    const addToContactsSheet = addToSheet(contactsListSheet); // curried function
    // https://developers.google.com/people/api/rest/v1/contactGroups/list
    const contactGroupResponses = getContactGroups(
      quotaUser,
      scriptProperties.getProperty("RESOURCE_NAME_ACTIVE"),
      scriptProperties?.getProperty("RESOURCE_NAME_GUEST"),
      scriptProperties?.getProperty("RESOURCE_NAME_STUDENT"),
      scriptProperties?.getProperty("RESOURCE_NAME_INACTIVE")
    );
    const [actives, guests, students, inactives] = getPeopleResponses(
      contactGroupResponses,
      quotaUser
    );
    const updateNeeded =
      forceUpdate === "forceUpdate"
        ? true
        : personnelChanges(
            ssLastUpdated,
            scriptProperties,
            actives,
            guests,
            inactives,
            students
          );
    let ssActiveData: Array<Array<string | undefined>>;
    let ssGuestData: Array<Array<string | undefined>>;
    let ssStudentData: Array<Array<string | undefined>>;
    let ssInactiveData: Array<Array<string | undefined>>;
    let rowCount = 2; // row 1 is the header
    const dt = new Date();

    if (updateNeeded) {
      ssActiveData = getContactsData("Active", actives);
      ssGuestData = getContactsData("Guest", guests);
      ssStudentData = getContactsData("Inactive", inactives);
      ssInactiveData = getContactsData("Students", students);

      if (ssActiveData.length > 0) {
        if (contactsListSheet.getLastRow() > 1) {
          contactsListSheet
            .getRange(2, 1, contactsListSheet.getLastRow() - 1, 7)
            .clearContent();
        }

        rowCount = addToContactsSheet(ssActiveData, rowCount);
        rowCount = addToContactsSheet(ssGuestData, rowCount);
        rowCount = addToContactsSheet(ssStudentData, rowCount);
        addToContactsSheet(ssInactiveData, rowCount);
        activeSpreadsheet.rename(`Sutherland Contacts ${dt.getFullYear()}`);
      }
    }

    return;
  }

  return { main };
})();
