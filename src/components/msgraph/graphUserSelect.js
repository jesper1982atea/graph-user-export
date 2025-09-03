// Canonical list of Graph user properties we select everywhere we fetch users
// Must use Graph property names (camelCase). Include displayName explicitly.
export const GRAPH_USER_SELECT_FIELDS = [
  'id', 'displayName', 'mail', 'userPrincipalName', 'jobTitle', 'department', 'companyName',
  'givenName', 'surname', 'mobilePhone', 'businessPhones', 'officeLocation', 'city', 'country', 'state',
  'streetAddress', 'postalCode', 'physicalDeliveryOfficeName', 'employeeId',
  'onPremisesDistinguishedName', 'onPremisesDomainName', 'onPremisesUserPrincipalName', 'onPremisesSamAccountName',
  // Navigation props like manager/managedDevices require $expand for details; we include names for CSV consistency
  'manager', 'managedDevices', 'assignedPlans', 'provisionedPlans', 'onPremisesExtensionAttributes'
];
