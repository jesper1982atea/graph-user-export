// Simple shared cache for chat members across components
export const chatMembersCache = new Map(); // chatId -> members array

export function getCachedChatMembers(chatId) {
  return chatMembersCache.get(chatId);
}

export async function fetchChatMembers(chatId, authHeader) {
  if (!chatId || !authHeader) return [];
  try {
    const res = await fetch(`https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(chatId)}/members`, {
      headers: { Authorization: authHeader },
    });
    if (!res.ok) return [];
    const data = await res.json();
    const members = Array.isArray(data.value) ? data.value : [];
    chatMembersCache.set(chatId, members);
    return members;
  } catch {
    return [];
  }
}
