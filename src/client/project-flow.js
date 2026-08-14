import {
  PROJECT_FLOW_PHASES
} from './config.js';

export function getProjectFlowTasksByPhase(tasks){
  const grouped = new Map(PROJECT_FLOW_PHASES.map(phase=>[phase.id, []]));
  (Array.isArray(tasks) ? tasks : []).forEach(task=>{
    const phaseId = PROJECT_FLOW_PHASES.some(phase=>phase.id === task.phaseId) ? task.phaseId : PROJECT_FLOW_PHASES[0].id;
    const list = grouped.get(phaseId) || [];
    list.push(task);
    grouped.set(phaseId, list);
  });
  grouped.forEach(list=>{
    list.sort((a, b)=>String(a.startDate || '').localeCompare(String(b.startDate || ''), 'no'));
  });
  return grouped;
}

export function formatProjectFlowStatusHours(milliseconds){
  return `${Math.max(1, Math.floor(milliseconds / 3600000))} t`;
}

export function formatProjectFlowStatusDays(milliseconds){
  return `${Math.max(1, Math.floor(milliseconds / 86400000))} d`;
}

export function formatProjectFlowStatusAge(milliseconds){
  return milliseconds >= 24 * 60 * 60 * 1000
    ? formatProjectFlowStatusDays(milliseconds)
    : formatProjectFlowStatusHours(milliseconds);
}

export function getProjectFlowTaskActivityTime(task, parseProjectFlowDate){
  const explicit = new Date(task?.updatedAt || task?.createdAt || 0).getTime();
  if (Number.isFinite(explicit) && explicit > 0) return explicit;
  const end = parseProjectFlowDate(task?.endDate);
  if (end) return end.getTime();
  const start = parseProjectFlowDate(task?.startDate);
  return start ? start.getTime() : 0;
}

export function getProjectFlowStatusForProject(project, tasks, helpers = {}){
  const parseProjectFlowDate = helpers.parseProjectFlowDate || (()=>null);
  const tasksByPhase = getProjectFlowTasksByPhase(tasks);
  const firstPhaseTasks = tasksByPhase.get(PROJECT_FLOW_PHASES[0].id) || [];
  if (!firstPhaseTasks.some(task=>task.completed)){
    const projectTime = new Date(project?.createdAt || 0).getTime();
    const untreatedAge = Number.isFinite(projectTime) && projectTime > 0 ? Date.now() - projectTime : 0;
    const stale = untreatedAge >= 24 * 60 * 60 * 1000;
    return {
      label: 'Ubehandlet',
      tone: stale ? 'danger' : 'idle',
      ageText: stale ? formatProjectFlowStatusAge(untreatedAge) : '',
      actionableAt: stale ? projectTime + (24 * 60 * 60 * 1000) : 0
    };
  }
  let lastCompletedPhase = PROJECT_FLOW_PHASES[0];
  for (const phase of PROJECT_FLOW_PHASES.slice(1)){
    const phaseTasks = tasksByPhase.get(phase.id) || [];
    if (!phaseTasks.length) break;
    const hasOpenTask = phaseTasks.some(task=>!task.completed);
    if (hasOpenTask){
      const latestActivity = Math.max(...phaseTasks.filter(task=>!task.completed).map(task=>getProjectFlowTaskActivityTime(task, parseProjectFlowDate)));
      const age = Number.isFinite(latestActivity) && latestActivity > 0 ? Date.now() - latestActivity : 0;
      const stale = age >= 7 * 24 * 60 * 60 * 1000;
      return {
        label: phase.label,
        tone: stale ? 'danger' : 'warning',
        ageText: stale ? formatProjectFlowStatusDays(age) : '',
        actionableAt: stale ? latestActivity + (7 * 24 * 60 * 60 * 1000) : 0
      };
    }
    lastCompletedPhase = phase;
  }
  return {
    label: lastCompletedPhase.label,
    tone: 'success'
  };
}
