const estado = document.querySelector("#estado");
const detalle = document.querySelector("#detalle-evento");

async function cargarDetalle() {
    const id = new URLSearchParams(window.location.search).get("id");
    if(!id){
        estado.textContent = "Elige un concierto desde la agenda.";
        return;
    }
    const respuesta = await fetch("../data/eventos.json");
     if (!respuesta.ok) {
    throw new Error(`No se pudo cargar la agenda. Código ${respuesta.status}`);
  }

  const agenda = await respuesta.json();
    // 3. Buscamos el concierto dentro de cada subcategoría.
  for (const categoria of agenda.categorias) {
    for (const subcategoria of categoria.subcategorias) {
      const evento = subcategoria.eventos.find(evento => evento.id === id);

      if (evento) {
        document.title = `${evento.titulo} | Agenda de conciertos`;

        document.querySelector("#genero").textContent =
          `${categoria.nombre} · ${subcategoria.nombre}`;

        document.querySelector("#titulo").textContent = evento.titulo;
        document.querySelector("#artista").textContent = evento.artista;
        document.querySelector("#descripcion").textContent = evento.descripcion;

        const fechaHora = document.querySelector("#fecha-hora");
        fechaHora.dateTime = `${evento.fecha}T${evento.hora}`;
        fechaHora.textContent = `${evento.fecha} · ${evento.hora}`;

        document.querySelector("#lugar").textContent = evento.lugar;

        document.querySelector("#entrada").textContent =
          evento.precioBs === 0 ? "Entrada libre" : `Entrada: Bs ${evento.precioBs}`;

        estado.textContent = "";
        detalle.hidden = false;
        return;
      }
    }
  }

  estado.textContent = "No se encontró el concierto solicitado.";

}
cargarDetalle().catch(error => {
  estado.textContent = error.message;
});
