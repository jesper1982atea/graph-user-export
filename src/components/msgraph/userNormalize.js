export function normalizeUser(u) {
  if (!u || typeof u !== 'object') return u;
  const ext = u.onPremisesExtensionAttributes || {};
  const flatExt = {};
  for (let i = 1; i <= 15; i++) {
    const k = `extensionAttribute${i}`;
    flatExt[k] = ext?.[k] ?? '';
  }
  const flatMgr = {
    managerDisplayName: u?.manager?.displayName || u?.managerDisplayName || '',
    managerMail: u?.manager?.mail || u?.managerMail || '',
    managerUserPrincipalName: u?.manager?.userPrincipalName || u?.managerUserPrincipalName || '',
    managerJobTitle: u?.manager?.jobTitle || u?.managerJobTitle || '',
  };
  // Derived summaries for common array-like attributes
  const assignedPlans = Array.isArray(u.assignedPlans) ? u.assignedPlans : [];
  const provisionedPlans = Array.isArray(u.provisionedPlans) ? u.provisionedPlans : [];
  const managedDevices = Array.isArray(u.managedDevices) ? u.managedDevices : [];
  const flatDerived = {
    assignedPlansCount: assignedPlans.length || 0,
    assignedPlansServices: assignedPlans.map(p => p.service || p.servicePlanId || '').filter(Boolean).join(', '),
    provisionedPlansCount: provisionedPlans.length || 0,
    provisionedPlansServices: provisionedPlans.map(p => p.service || p.servicePlanId || '').filter(Boolean).join(', '),
    managedDevicesCount: managedDevices.length || 0,
  };
  return { ...u, ...flatExt, ...flatMgr, ...flatDerived };
}
