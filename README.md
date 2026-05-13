<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Agenda de Trabajo Dinámica Pro | Mercantil Seguros</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        .glass-card {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.2);
        }
        .status-badge {
            transition: all 0.3s ease;
        }
        input::placeholder, textarea::placeholder {
            color: #cbd5e1;
        }
    </style>
</head>
<body class="bg-slate-50 min-h-screen font-sans text-slate-900">

    <nav class="bg-indigo-800 text-white shadow-lg p-4 sticky top-0 z-50">
        <div class="max-w-6xl mx-auto flex justify-between items-center">
            <h1 class="text-xl font-bold flex items-center gap-2">
                <i class="fas fa-business-time"></i> Agenda de Ejecución
            </h1>
            <div id="date-display" class="hidden md:block text-sm font-medium opacity-90"></div>
            <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition-all shadow-md">
                <i class="fas fa-file-excel"></i> Exportar a Excel
            </button>
        </div>
    </nav>

    <main class="max-w-6xl mx-auto p-4 md:p-6 space-y-6">
        
        <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div class="glass-card p-4 rounded-xl shadow-sm border-l-4 border-indigo-500">
                <p class="text-[10px] md:text-xs text-slate-500 uppercase font-bold tracking-wider">Total Visitas</p>
                <h2 id="total-count" class="text-xl md:text-2xl font-bold text-slate-800">0</h2>
            </div>
            <div class="glass-card p-4 rounded-xl shadow-sm border-l-4 border-emerald-500">
                <p class="text-[10px] md:text-xs text-slate-500 uppercase font-bold tracking-wider">Completadas</p>
                <h2 id="completed-count" class="text-xl md:text-2xl font-bold text-slate-800">0</h2>
            </div>
            <div class="glass-card p-4 rounded-xl shadow-sm border-l-4 border-amber-500">
                <p class="text-[10px] md:text-xs text-slate-500 uppercase font-bold tracking-wider">Pendientes</p>
                <h2 id="pending-count" class="text-xl md:text-2xl font-bold text-slate-800">0</h2>
            </div>
            <div class="glass-card p-4 rounded-xl shadow-sm border-l-4 border-blue-400">
                <p class="text-[10px] md:text-xs text-slate-500 uppercase font-bold tracking-wider">Tiempo Estimado</p>
                <h2 id="total-time" class="text-xl md:text-2xl font-bold text-slate-800">0 <span class="text-sm font-normal">min</span></h2>
            </div>
        </div>

        <section class="glass-card p-6 rounded-2xl shadow-md border border-slate-200">
            <h3 class="text-lg font-semibold mb-4 flex items-center gap-2 text-indigo-700">
                <i class="fas fa-calendar-plus"></i> Programar Ejecución
            </h3>
            <form id="visit-form" class="space-y-4">
                <div class="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 gap-4">
                    <div class="md:col-span-2 lg:col-span-1">
                        <label class="block text-xs font-bold text-slate-600 mb-1">Cliente / Empresa</label>
                        <input type="text" id="client-name" required placeholder="Nombre del cliente" 
                            class="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">Día y Hora</label>
                        <input type="datetime-local" id="visit-datetime" required
                            class="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">Duración Estimada (min)</label>
                        <input type="number" id="visit-duration" required placeholder="Ej: 45" min="5" step="5"
                            class="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-600 mb-1">Objetivo</label>
                        <select id="visit-objective" class="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm">
                            <option value="Prospección">Prospección</option>
                            <option value="Presentación">Presentación</option>
                            <option value="Renovación">Renovación</option>
                            <option value="Cierre">Cierre</option>
                            <option value="Técnica">Visita Técnica</option>
                        </select>
                    </div>
                </div>
                
                <div class="grid grid-cols-1 md:grid-cols-4 gap-4">
                    <div class="md:col-span-1">
                        <label class="block text-xs font-bold text-slate-600 mb-1">Prioridad</label>
                        <select id="visit-priority" class="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm">
                            <option value="Alta">Alta</option>
                            <option value="Media" selected>Media</option>
                            <option value="Baja">Baja</option>
                        </select>
                    </div>
                    <div class="md:col-span-3">
                        <label class="block text-xs font-bold text-slate-600 mb-1">Notas u Observaciones</label>
                        <textarea id="visit-notes" rows="1" placeholder="Detalles adicionales, requisitos, puntos a tratar..." 
                            class="w-full p-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none text-sm resize-none"></textarea>
                    </div>
                </div>

                <div class="flex justify-end">
                    <button type="submit" class="w-full md:w-auto bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2.5 px-8 rounded-lg transition-all flex items-center justify-center gap-2 shadow-sm">
                        <i class="fas fa-plus"></i> Agregar a mi Ruta
                    </button>
                </div>
            </form>
        </section>

        <section class="space-y-4">
            <div class="flex justify-between items-center px-2">
                <h3 class="text-lg font-semibold text-slate-700">Hoja de Ruta</h3>
                <div class="flex gap-4">
                    <button onclick="clearCompleted()" class="text-sm text-slate-500 hover:text-red-500 transition-colors">
                        <i class="fas fa-trash-alt"></i> Limpiar completadas
                    </button>
                </div>
            </div>
            
            <div id="visits-list" class="space-y-3">
                </div>

            <div id="empty-state" class="hidden text-center py-16 bg-white rounded-2xl border border-dashed border-slate-300 text-slate-400">
                <i class="fas fa-route text-5xl mb-4 opacity-20"></i>
                <p class="font-medium">Tu hoja de ruta está vacía.</p>
                <p class="text-sm">Comienza agregando tu primera visita arriba.</p>
            </div>
        </section>
    </main>

    <script>
        let visits = JSON.parse(localStorage.getItem('workPlanVisitsPro_v2')) || [];

        const visitForm = document.getElementById('visit-form');
        const visitsList = document.getElementById('visits-list');
        const emptyState = document.getElementById('empty-state');
        
        function updateDate() {
            const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' };
            const display = document.getElementById('date-display');
            if(display) display.innerText = new Date().toLocaleDateString('es-ES', options);
        }
        setInterval(updateDate, 1000);
        updateDate();

        function saveVisits() {
            localStorage.setItem('workPlanVisitsPro_v2', JSON.stringify(visits));
            renderVisits();
            updateStats();
        }

        function updateStats() {
            const completed = visits.filter(v => v.done);
            const pending = visits.filter(v => !v.done);
            
            document.getElementById('total-count').innerText = visits.length;
            document.getElementById('completed-count').innerText = completed.length;
            document.getElementById('pending-count').innerText = pending.length;
            
            const totalMinutes = pending.reduce((acc, curr) => acc + parseInt(curr.duration || 0), 0);
            document.getElementById('total-time').innerHTML = `${totalMinutes} <span class="text-sm font-normal">min</span>`;
        }

        function toggleVisit(id) {
            visits = visits.map(v => v.id === id ? {...v, done: !v.done} : v);
            saveVisits();
        }

        function deleteVisit(id) {
            if(confirm('¿Estás seguro de eliminar esta visita?')) {
                visits = visits.filter(v => v.id !== id);
                saveVisits();
            }
        }

        function clearCompleted() {
            if(confirm('Se eliminarán todas las visitas marcadas como completadas. ¿Continuar?')) {
                visits = visits.filter(v => !v.done);
                saveVisits();
            }
        }

        visitForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const newVisit = {
                id: Date.now(),
                client: document.getElementById('client-name').value,
                datetime: document.getElementById('visit-datetime').value,
                duration: document.getElementById('visit-duration').value,
                objective: document.getElementById('visit-objective').value,
                priority: document.getElementById('visit-priority').value,
                notes: document.getElementById('visit-notes').value,
                done: false
            };
            visits.push(newVisit);
            visitForm.reset();
            saveVisits();
        });

        function formatVisitTime(dateStr) {
            const d = new Date(dateStr);
            return d.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });
        }

        function renderVisits() {
            visitsList.innerHTML = '';
            
            if (visits.length === 0) {
                emptyState.classList.remove('hidden');
                return;
            } else {
                emptyState.classList.add('hidden');
            }

            const sortedVisits = [...visits].sort((a, b) => {
                if (a.done !== b.done) return a.done ? 1 : -1;
                return new Date(a.datetime) - new Date(b.datetime);
            });

            sortedVisits.forEach(visit => {
                const priorityStyles = {
                    "Alta": "text-red-600 bg-red-50 border-red-100",
                    "Media": "text-amber-600 bg-amber-50 border-amber-100",
                    "Baja": "text-emerald-600 bg-emerald-50 border-emerald-100"
                };

                const card = document.createElement('div');
                card.className = `glass-card p-4 rounded-xl shadow-sm border transition-all flex flex-col gap-3 ${visit.done ? 'opacity-60 bg-slate-50 grayscale' : 'hover:shadow-md border-slate-200'}`;
                
                card.innerHTML = `
                    <div class="flex flex-col md:flex-row md:items-center justify-between gap-4">
                        <div class="flex items-center gap-4 flex-1">
                            <button onclick="toggleVisit(${visit.id})" class="text-3xl flex-shrink-0 ${visit.done ? 'text-emerald-500' : 'text-slate-300 hover:text-indigo-400'}">
                                <i class="${visit.done ? 'fas fa-check-circle' : 'far fa-circle'}"></i>
                            </button>
                            <div class="flex-1">
                                <div class="flex items-center gap-2 mb-1">
                                    <h4 class="font-bold text-lg ${visit.done ? 'line-through text-slate-500' : 'text-slate-800'}">${visit.client}</h4>
                                    <span class="text-[10px] px-2 py-0.5 rounded-full border font-bold uppercase ${priorityStyles[visit.priority]}">
                                        ${visit.priority}
                                    </span>
                                </div>
                                <div class="flex flex-wrap items-center gap-y-1 gap-x-4 text-sm text-slate-600">
                                    <span class="flex items-center gap-1.5 font-medium">
                                        <i class="far fa-calendar-alt text-indigo-500"></i> ${formatVisitTime(visit.datetime)}
                                    </span>
                                    <span class="flex items-center gap-1.5">
                                        <i class="far fa-clock text-indigo-500"></i> ${visit.duration} min
                                    </span>
                                    <span class="flex items-center gap-1.5">
                                        <i class="fas fa-tag text-slate-400"></i> ${visit.objective}
                                    </span>
                                </div>
                            </div>
                        </div>
                        <div class="flex items-center justify-end md:border-l md:pl-4">
                            <button onclick="deleteVisit(${visit.id})" class="p-2 text-slate-400 hover:text-red-500 transition-colors">
                                <i class="fas fa-trash-alt"></i>
                            </button>
                        </div>
                    </div>
                    ${visit.notes ? `
                    <div class="mt-1 p-3 bg-slate-100/50 rounded-lg border border-slate-100 text-sm text-slate-700 italic">
                        <i class="fas fa-sticky-note text-indigo-300 mr-2"></i> ${visit.notes}
                    </div>` : ''}
                `;
                visitsList.appendChild(card);
            });
        }

        function exportToExcel() {
            if (visits.length === 0) {
                alert('No hay datos para exportar.');
                return;
            }

            const dataToExport = visits.map(v => ({
                'Cliente': v.client,
                'Fecha y Hora': v.datetime.replace('T', ' '),
                'Duración (min)': v.duration,
                'Objetivo': v.objective,
                'Prioridad': v.priority,
                'Estado': v.done ? 'Completado' : 'Pendiente',
                'Notas': v.notes || ''
            }));

            const worksheet = XLSX.utils.json_to_sheet(dataToExport);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Agenda de Visitas");

            const wscols = [
                {wch: 30}, {wch: 20}, {wch: 15}, {wch: 20}, {wch: 10}, {wch: 15}, {wch: 50}
            ];
            worksheet['!cols'] = wscols;

            XLSX.writeFile(workbook, `Agenda_Trabajo_${new Date().toISOString().split('T')[0]}.xlsx`);
        }

        renderVisits();
        updateStats();
    </script>
</body>
</html>
