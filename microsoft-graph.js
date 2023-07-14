// Realiza una solicitud GET a Microsoft Graph
function obtenerDatosMicrosoftGraph() {
    // Realiza la solicitud GET para obtener los sitios activos
    fetch('https://graph.microsoft.com/v1.0/sites?search=*.sharepoint.com', {
      headers: {
        'Authorization': 'Bearer {tu-token-de-acceso}'
      }
    })
    .then(response => response.json())
    .then(data => {
      // Procesa la respuesta y obtén los sitios activos
      const sitios = data.value;
  
      // Recorre los sitios y realiza solicitudes adicionales
      sitios.forEach(sitio => {
        // Realiza la solicitud GET para obtener la biblioteca de documentos raíz del sitio
        fetch(`https://graph.microsoft.com/v1.0/sites/${sitio.id}/drive/root`, {
          headers: {
            'Authorization': 'Bearer {tu-token-de-acceso}'
          }
        })
        .then(response => response.json())
        .then(data => {
          // Procesa la respuesta y muestra los datos de la biblioteca de documentos
          const biblioteca = data.name;
          console.log(`Biblioteca de documentos en ${sitio.displayName}: ${biblioteca}`);
        });
  
        // Realiza la solicitud GET para obtener las listas del sitio
        fetch(`https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists`, {
          headers: {
            'Authorization': 'Bearer {tu-token-de-acceso}'
          }
        })
        .then(response => response.json())
        .then(data => {
          // Procesa la respuesta y muestra los nombres de las listas
          const listas = data.value.map(lista => lista.name);
          console.log(`Listas en ${sitio.displayName}: ${listas.join(', ')}`);
        });
      });
    });
  }
  
  // Llama a la función para obtener los datos de Microsoft Graph
  obtenerDatosMicrosoftGraph();
