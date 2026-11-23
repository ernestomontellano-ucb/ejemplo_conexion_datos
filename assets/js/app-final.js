const listaEventos = document.querySelector(".lista_eventos");

async function cargarEventos() {
    const respuesta = await fetch("./data/eventos.json");
    if(!respuesta.ok){
        throw new Error(`No se puede cargar la agenda. Código ${respuesta.status}`);
    }
    const agenda = await respuesta.json();
    for(const categoria of agenda.categorias){
        for(const subcategoria of categoria.subcategorias){
            for(const evento of subcategoria.eventos){
                const card = document.createElement("article");
                card.className = "card-evento";
                card.id = evento.id;

                const entrada = evento.precioBs === 0
                ? "Entrada libre"
                : `Bs ${evento.precioBs}`;

                card.innerHTML = `
                    <p class="categoria-evento">${categoria.nombre} - ${subcategoria.nombre}</p>
                    <h3>${evento.titulo}</h3>
                    <p>${evento.artista}</p>
                    <p class="dato-evento">
                        <i class="ph ph-calendar-dots" aria-hidden="true"></i>
                        <time datetime="${evento.fecha}T${evento.hora}">${evento.fecha} - ${evento.hora}</time>
                    </p>
                    <p class="dato-evento">
                        <i class="ph ph-map-pin" aria-hidden="true"></i>
                        <span>${evento.lugar}</span>
                    </p>
                    <p class="dato-evento precio-evento">
                        <i class="ph ph-ticket" aria-hidden="true"></i>
                        <span>${entrada}</span>
                    </p>
                    <a href="./pages/evento.html?id=${encodeURIComponent(evento.id)}">Ver detalles del concierto</a>
                `;
                listaEventos.append(card);
            }
        }
    }
}

cargarEventos().catch(error =>{
    listaEventos.textContent = error.message;
})
