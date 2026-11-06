// Uses the timeline's action attributes and handlers, so both views save the same tasks.
export function renderProjectFlowList(root, projects, tasks, helpers){
  const list = document.createElement('div');
  list.className = 'project-flow-list';
  for (const project of projects){
    const section = document.createElement('section');
    section.className = 'project-flow-list-project';
    const heading = document.createElement('h3');
    heading.textContent = helpers.getProjectLabel(project);
    section.append(heading);
    const projectTasks = tasks.filter(task=>task.projectId === project.id);
    if (!projectTasks.length){
      const empty = document.createElement('p');
      empty.className = 'muted-text';
      empty.textContent = 'Ingen oppgaver. Bruk «Ny oppgave» for å legge til.';
      section.append(empty);
    }
    for (const phase of helpers.phases){
      const phaseTasks = projectTasks.filter(task=>task.phaseId === phase.id)
        .sort((a, b)=>String(a.startDate || a.date || '').localeCompare(String(b.startDate || b.date || '')));
      if (!phaseTasks.length) continue;
      const group = document.createElement('div');
      group.className = 'project-flow-list-phase';
      const phaseTitle = document.createElement('h4');
      phaseTitle.textContent = phase.label;
      group.append(phaseTitle);
      for (const task of phaseTasks){
        const card = document.createElement('article');
        card.className = 'project-flow-list-task';
        card.classList.toggle('is-completed', Boolean(task.completed));
        const label = helpers.getTaskLabel(task);
        const checkLabel = document.createElement('label');
        checkLabel.className = 'project-flow-list-check';
        const check = document.createElement('input');
        check.type = 'checkbox';
        check.checked = Boolean(task.completed);
        check.dataset.projectFlowToggle = task.id;
        check.dataset.projectId = project.id;
        check.setAttribute('aria-label', `Fullført: ${label}`);
        checkLabel.append(check);
        const body = document.createElement('div');
        const title = document.createElement('strong');
        title.textContent = label;
        const dates = document.createElement('p');
        dates.className = 'muted-text';
        dates.textContent = `${helpers.formatDate(task.startDate || task.date)} – ${helpers.formatDate(task.endDate || task.date)}`;
        const responsible = document.createElement('p');
        responsible.className = 'muted-text';
        responsible.textContent = helpers.getResponsibleLabel(project);
        body.append(title, dates, responsible);
        const edit = document.createElement('button');
        edit.type = 'button';
        edit.className = 'btn alt';
        edit.textContent = 'Rediger';
        edit.dataset.projectFlowEdit = task.id;
        edit.dataset.projectId = project.id;
        edit.setAttribute('aria-label', `Rediger ${label}`);
        card.append(checkLabel, body, edit);
        group.append(card);
      }
      section.append(group);
    }
    list.append(section);
  }
  root.append(list);
}
