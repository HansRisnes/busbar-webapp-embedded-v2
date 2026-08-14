export function readLocalText(key, fallback = ''){
  if (!key || typeof localStorage === 'undefined') return fallback;
  try{
    const value = localStorage.getItem(key);
    return value == null ? fallback : String(value);
  }catch(_err){
    return fallback;
  }
}

export function writeLocalText(key, value){
  if (!key || typeof localStorage === 'undefined') return false;
  try{
    localStorage.setItem(key, String(value ?? ''));
    return true;
  }catch(_err){
    return false;
  }
}

export function removeLocalItem(key){
  if (!key || typeof localStorage === 'undefined') return false;
  try{
    localStorage.removeItem(key);
    return true;
  }catch(_err){
    return false;
  }
}

export function hasLocalItem(key){
  if (!key || typeof localStorage === 'undefined') return false;
  try{
    return localStorage.getItem(key) != null;
  }catch(_err){
    return false;
  }
}

export function readLocalJson(key, fallback = null){
  const raw = readLocalText(key, '');
  if (!raw) return fallback;
  try{
    return JSON.parse(raw);
  }catch(_err){
    return fallback;
  }
}

export function writeLocalJson(key, value){
  if (!key || typeof localStorage === 'undefined') return false;
  try{
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  }catch(_err){
    return false;
  }
}

export function listLocalKeys(predicate = null){
  if (typeof localStorage === 'undefined') return [];
  const keys = [];
  try{
    for (let i = 0; i < localStorage.length; i += 1){
      const key = localStorage.key(i);
      if (!key) continue;
      if (!predicate || predicate(key)){
        keys.push(key);
      }
    }
  }catch(_err){}
  return keys;
}

export function readSessionJson(key, fallback = null){
  if (!key || typeof sessionStorage === 'undefined') return fallback;
  try{
    const raw = sessionStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  }catch(_err){
    return fallback;
  }
}

export function writeSessionJson(key, value){
  if (!key || typeof sessionStorage === 'undefined') return false;
  try{
    sessionStorage.setItem(key, JSON.stringify(value));
    return true;
  }catch(_err){
    return false;
  }
}

export function removeSessionItem(key){
  if (!key || typeof sessionStorage === 'undefined') return false;
  try{
    sessionStorage.removeItem(key);
    return true;
  }catch(_err){
    return false;
  }
}
