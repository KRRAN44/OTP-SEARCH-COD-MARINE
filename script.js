const archivo = document.getElementById("archivo");
const boton = document.getElementById("buscar");
const pregunta = document.getElementById("pregunta");
const respuesta = document.getElementById("respuesta");

let documentoTexto = "";

// Cargar documentos (TXT, Excel, Word)
archivo.addEventListener("change", async () => {

    const file = archivo.files[0];

    if(!file) return;


    // TXT
    if(file.name.endsWith(".txt")){

        const lector = new FileReader();

        lector.onload = function(e){

            documentoTexto = e.target.result;
            respuesta.innerHTML = "📄 TXT CARGADO ✅";

        };

        lector.readAsText(file);

    }


    // Excel (.xlsx / .xls)
    else if(
        file.name.endsWith(".xlsx") || 
        file.name.endsWith(".xls")
    ){

        const datos = await file.arrayBuffer();

        const libro = XLSX.read(datos);

        let texto = "";

        libro.SheetNames.forEach(nombre => {

            const hoja = libro.Sheets[nombre];

            texto += XLSX.utils.sheet_to_txt(hoja);

        });

        documentoTexto = texto;

        respuesta.innerHTML = "📊 EXCEL CARGADO ✅";

    }


    // Word (.docx)
    else if(file.name.endsWith(".docx")){

        const datos = await file.arrayBuffer();

        const resultado = await mammoth.extractRawText({
            arrayBuffer: datos
        });

        documentoTexto = resultado.value;

        respuesta.innerHTML = "📝 WORD CARGADO ✅";

    }


    else{

        respuesta.innerHTML = "❌ Formato no compatible";

    }

});

// Buscar información
boton.addEventListener("click", () => {

    if(documentoTexto === ""){
        respuesta.innerHTML = "⚠️ Primero carga un documento.";
        return;
    }

    const preguntaUsuario = pregunta.value.toLowerCase().trim();

    if(preguntaUsuario === ""){
        respuesta.innerHTML = "⚠️ Inserta los digitos, Ex: 4444 , 9999.";
        return;
    }

    const palabras = preguntaUsuario.split(/\s+/);

    const lineas = documentoTexto.split("\n");

    let encontrados = [];

    for(let linea of lineas){

        let texto = linea.toLowerCase();

        for(let palabra of palabras){

            if(texto.includes(palabra)){

                // Evita repetir la misma línea
                if(!encontrados.includes(linea)){
                    encontrados.push(linea);
                }

                break;
            }

        }

    }

    if(encontrados.length > 0){

        respuesta.innerHTML =
        "OTP'S <b>" + encontrados.length + "</b> resultado(s):<br><br>" +
        encontrados.join("<br><br>");

    }else{

        respuesta.innerHTML =
        "❌ No encontré información relacionada.";

    }

});