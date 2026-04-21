const activeTags = new Set();

function addTag(label) {
    if (activeTags.has(label)) return;
    activeTags.add(label);
    renderTags();
}

function removeTag(label) {
    activeTags.delete(label);
    renderTags();
}

function clearAllTags() {
    activeTags.clear();
    renderTags();
}

function renderTags() {
    const container = document.getElementById('tag-container');
    container.innerHTML = '';

    activeTags.forEach(label => {
        const tag = document.createElement('span');
        tag.className = 'inline-flex items-center gap-1.5 bg-primary/10 text-primary border border-primary/20 px-3 py-1.5 rounded-full text-xs font-bold';
        tag.innerHTML = `
            ${label}
            <button onclick="removeTag('${label}')"
                class="w-3.5 h-3.5 rounded-full bg-primary/20 hover:bg-primary hover:text-white flex items-center justify-center transition-colors text-[10px] leading-none">
                ✕
            </button>`;
        container.appendChild(tag);
    });

    if (activeTags.size > 1) {
        const btn = document.createElement('button');
        btn.className = 'text-primary text-xs font-bold ml-2 hover:underline';
        btn.textContent = 'Limpiar todo';
        btn.onclick = clearAllTags;
        container.appendChild(btn);
    }
}