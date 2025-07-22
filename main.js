
document.getElementById("sistema-select").addEventListener("change", async function () {
    const sistema = this.value;
    if (!sistema) return;
    const response = await fetch(`./data/${sistema}.json`);
    const data = await response.json();
    iniciarConversacion(data);
});

function iniciarConversacion(modulo) {
    const preguntas = modulo.modulos[0].preguntas;
    const codigos = modulo.modulos[0].codigos;
    const inicio = modulo.modulos[0].inicio;
    const respuestas = {};
    let actual = inicio;

    const chatBox = document.getElementById("chat-box");
    const controls = document.getElementById("controls");
    chatBox.innerHTML = "";
    controls.innerHTML = "";

    mostrarPregunta(actual);

    function mostrarPregunta(id) {
        const p = preguntas[id];
        const div = document.createElement("div");
        div.className = "chat-bubble";
        div.textContent = p.texto;
        chatBox.appendChild(div);
        controls.innerHTML = "";

        if (p.tipo === "opciones") {
            for (let [opcion, destino] of Object.entries(p.opciones)) {
                const btn = document.createElement("button");
                btn.textContent = opcion;
                btn.onclick = () => {
                    respuestas[id] = opcion;
                    if (destino.startsWith("01")) {
                        mostrarCodigo(destino);
                    } else if (destino === "fin") {
                        mostrarFin();
                    } else {
                        mostrarPregunta(destino);
                    }
                };
                controls.appendChild(btn);
            }
        }
    }

function mostrarCodigo(codigo) {
    const obj = codigos[codigo];
    let resultado;
    if (typeof obj === "object" && obj !== null) {
        resultado = `${codigo} (${obj.nivel}): ${obj.descripcion}`;
    } else {
        resultado = `${codigo}: ${obj || "Descripción no disponible"}`;
    }
    const div = document.createElement("div");
    div.className = "chat-bubble";
    div.innerHTML = `<strong>Código sugerido:</strong> ${resultado}`;
    chatBox.appendChild(div);
    controls.innerHTML = "";
}


    function mostrarFin() {
        const div = document.createElement("div");
        div.className = "chat-bubble";
        div.innerHTML = `<strong>No se pudo determinar el código con la información proporcionada.</strong>`;
        chatBox.appendChild(div);
        controls.innerHTML = "";
    }
}
