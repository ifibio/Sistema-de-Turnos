/*
   Copyright 2020 JJC
*/

function asistencia(e) {
  if (e.range.getSheet().getName() == 'Solicitud de Asistencia') { // me deja elegir con qué formulario debo disparar esta función.
    // como cada formulario es una hoja de la planilla, recupero el nombre de la planilla que se llena
    // Esta función dispara la creación del evento de asistenci al IFIBIO. 
    // va a parar al calendario cuyo ID sale de la pestaña Sector
    // Obtener la planilla
    var ss = SpreadsheetApp.getActiveSpreadsheet(); // en dónde se ejecuta el script, podría hardcodear el ID y evitar ejecuciones en cualquier lado
    // el ID de Actividades IFIBIO es: 
    var sheet = ss.getSheetByName('Solicitud de Asistencia'); // de dónde salen los datos, me sirve por si quiero cambiar de dónde saco los valores. Por eso le puse el nombre
    // Chequear si hay datos y obtener la última fila con valores   
    var valoresA = sheet.getRange("B1:B").getValues(); // todos los valores de la columna B, yo usaria una columna que si o si deba estar
    var last = valoresA.filter(String).length; // la ultima fila con valores de la columna B.- Hay alguna forma de hacerlo dinámico?

    /*
        Numeración de columnas
        
        A B C D E F G H I J  K  L  M  N  O  P  Q  R  S  T  U  V  W  X  Y  Z     En letras
        1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26    En números
        0 1 2 3 4 5 6 7 8 9  10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25     En arrays
                
        Obtener los datos del rango
          
        getRange(row, column, optNumRows, optNumColumns)
             
        row           --- entero --- primera fila en números, así que empieza en 1
        column        --- entero --- primera columna
        optNumRows    --- entero --- número de filas en el rango (hasta dónde en filas)
        optNumColumns --- entero --- número de columnas en el rango (hasta dónde en columnas)
    */
    // si agrando la cantidad de columnas a leer debo variar el último valor de la parte .getRange(last,1,1,ULTIMA_COLUMNA_PARA_LEER) 
    var data = sheet.getRange(last, 1, 1, 26).getValues(); // obtener valores (en crudo, si hay una fórmula medio que no funciona bien), encontré problemas con interpretación de fechas
    var array = sheet.getRange(last, 1, 1, 26).getDisplayValues(); // me sirve para leer contenidos que se muestran pero se calculan en el momento y es lo que vemos en la planilla

    /*  create variables 
          la variable se crea sacando la info con data[][] si viene directamente del formulario o array[][] si es resultado de otra cosa, como una fórmula
          el "0" acá representa la primera fila o columna a tomar en cuenta DENTRO del rango [FILA][COLUMNA], entonces contando desde 1 siempre es n+1
          
          si [n][n] es [0][4] = la primera fila (0+1) y la quinta columna (4+1)
    
    Al 1/12/2020 las columnas son
           0 - A          	1                   2                           3 - D         4                  5            6          
      Timestamp	Nombre Y Apellido	Sector Principal de Trabajo	Día de Asistencia	Email Address	Hora de ingreso	Hora de egreso	
      
                7                                  8 - I                                        9                     10          
      Se repite el mismo día y horario	¿Durante su tarea debe ingresar a otro sector 	Sector Secundario	Sector Terciario	
      
                 11             12             13 - N           14               15          	 16                 17				                                                                                                                                                                                                                                                            
      Sector Cuaternario	sector interno   Referente	  Calendario ID	  Diferencia con HOY	Color		Dia de la semana																						
     
           18         19             20             21 - V  
      AUTORIZADO	Estado    "columna vacia"     ID evento
    
    */

    var usuario = data[0][1]; //nombre del que peticiona
    var sector = data[0][2]; //lugar al que va
    var dia = array[0][3]; // fecha
    // var dia = data[0][3]; // fecha
    var mail = data[0][4]; // mail de registro
    var ingreso = dia + ' ' + array[0][5] + ' GMT-0300'; // hora de entrada el GMT corresponde a Peronia, las comillas simples se usan para poner texto
    var salida = dia + ' ' + array[0][6] + ' GMT-0300'; // hora de salida
    var frecuente = array[0][7]; // Se va a repetir una y otra vez cada semana?
    var sectorsec = array[0][9];
    var sectorter = array[0][10];
    var sectorcua = array[0][11];
    var otrosector = array[0][9] + ' + ' + array[0][10] + ' + ' + array[0][11];
    var interno = array[0][12];
    var referente = array[0][13]; // de quién depende el lugar
    var idcal = array[0][14]; // el ID del calendario
    var colorevent = array[0][16]; // el color del evento, parece que hay un error en esa config
    var diasemana = array[0][17]; // dia de la semana
    var diasemanaObjeto = CalendarApp.Weekday[diasemana]; // crear el Objeto que indica el día de la semana
    var autorizado = array[0][18]; // esta autorizado?

    // EL calendario de cada lugar
    var calendar = CalendarApp.getCalendarById(idcal); // lo hice flexible así se puede modificar en la pestaña de Sectores pero me mantengo trabajando en Solicitud

    /* 
  Un evento sencillo se puede crear con
      
    calendar.createEvent(usuario, new Date(ingreso), new Date(salida),   {location:sector, guests:mail});
           
     Hay tres condiciones para crear un evento que vienen de la columna "Se repite cada semana el mismo día y horario" cuando tiene valor 'Sí' o 'No'
     1- Asistencia por una sola vez = 'No': se hace con createEvent(title, startTime, endTime, options)
     2- Asistencia semanal (bioterio? mantenimiento?) = 'Semanal' con la clase addWeeklyRule() : con createEventSeries(title, startTime, endTime, recurrence, options) y newRecurrence().addWeeklyRule() o newRecurrence().addMonthlyRule()
     3- Asistencia mensual = 'Semanal' : con createEventSeries(title, startTime, endTime, recurrence, options) y newRecurrence().addMonthlyRule()
     
     El tema es que hay que distinguir entre repetición en la misma fecha (ejemplo, todos los 11) o en el mismo día (todos los martes)
     Si se repite una vez por mes los martes, por ejemplo, hay 4 martes por mes
     
     si se le pone .time() se puede decidir el límite, por ejemplo si sólo queremos que esto se repita como máximo 3 semanas
     
     onlyOnWeekday(CalendarApp.Weekday.WEDNESDAY) 
     
     .addWeeklyRule().onlyOnWeekday(CalendarApp.Weekday.WEDNESDAY) 
     
   Un evento lo podría describir como
    
     var event = {
      summary: usuario+' en '+sector,
      location: sector,
      description: 'Sector Interno: '+interno+' Otros Sectores:  '+otrosector ,
      }
    */

    // ======== SECCION DE AVISOS =========

    // == verificador de fechas ==
    // el horario de entrada debe ser menor al horario de salida, por lo tanto salida - entrada > 0
    // pongo un condicional para verificar las fechas
    var entra = new Date(ingreso);
    var sale = new Date(salida);
    var diffDate = new Date(sale - entra); // la expresion matematica que define la diferencia de fechas
    if (diffDate > 0) {
      var entrada_salida = 'correcto'; // le doy el valor  
    }
    else {
      var entrada_salida = 'incorrecto'; // está mal y se deberá resolver más adelante
    }

    // == termina el verificador de fechas ==

    // == con esto aviso si es semanal, mensual o no ==

     if (frecuente == "Semanal") {
      var frecuenciamensaje = '. Con frecuencia SEMANAL.';
      }
      else if (frecuente == "Mensual") {
      var frecuenciamensaje = '. Con frecuencia MENSUAL.';
      }
      else {
      var frecuenciamensaje = '';
      }

    // ahora bien, voy a crear un if muy grande que involucra a todo lo que dispara los mails. Me conviene?

    if (entrada_salida == 'correcto') {


      // == verificador de eventos en el mismo día ==
      // verifico que existan asistencias en todo el día más las asistencias que se superpongan

      var conflicts = calendar.getEventsForDay(new Date(dia)); // obtiene el número de turnos en ese día
      //var seisHoras = new Date(ingreso + (6 * 60 * 60 * 1000));
      //var franja = CalendarApp.getCalendarById(idcal).getEvents(new Date(ingreso), new Date(ingreso + (6 * 60 * 60 * 1000)));  // los turnos de seis horas que se superponen
      // == fin verificador de eventos en el mismo día ==

      // Filtrado de autorización
      if (autorizado == "Yes" || autorizado == "Se" || autorizado == "Si" || autorizado == "S" || autorizado == "s" || autorizado == "SI" || autorizado == "si" || autorizado == "sí" || autorizado == "SÍ" || autorizado == "Sí" || autorizado == "Y") { // Me fijo si está autorizado y dispara el pedido

        if (frecuente == "Semanal") { // una vez por semana el mismo día, ejemplo Lunes
          var recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(CalendarApp.Weekday[diasemanaObjeto]); // REGLA SEMANAL
          var event = calendar.createEventSeries(usuario, new Date(ingreso), new Date(salida), recurrence, { location: sector, guests: mail, sendNotifications: true }); // se le agrega la propiedad de repetir cada semana, 
          var turnoid = event.getId(); // el ID del evento
          var titulo = event.setTitle(usuario + ' ' + sector);
          //var colorear = event.setColor(colorevent); // le doy color, no funcionó 
          var descripcion = event.setDescription('Interno: ' + interno + ' Otros: ' + otrosector); // agrego una descripcion, me sirve para buscar cosas
          var modificar = event.setGuestsCanModify(true) // los invitados pueden cambiar ESE evento y nada mas
          var idcelda = sheet.getRange(last, 22, 1, 1).setValue(turnoid); // permite escribir el valor del id en una celda
        }

        else if (frecuente == "Mensual") { // una vez por mes, ejemplo el segundo martes.
          var recurrence = CalendarApp.newRecurrence().addMonthlyRule().onlyOnWeekday(CalendarApp.Weekday[diasemanaObjeto]).onlyOnMonthDays([1, 2, 3, 4, 5, 6, 7]); // REGLA MENSUAL el primer DIA del mes que coincide, en este caso es siempre la primera semana del mes
          var event = calendar.createEventSeries(usuario, new Date(ingreso), new Date(salida), recurrence, { location: sector, guests: mail, sendNotifications: true }); // se le agrega la propiedad de repetir cada mes
          var turnoid = event.getId();
          var titulo = event.setTitle(usuario + ' ' + sector);
          //var colorear = event.setColor(colorevent); // le doy color
          var descripcion = event.setDescription('Interno: ' + interno + ' Otros: ' + otrosector); // agrego una descripcion, me sirve para buscar cosas
          var modificar = event.setGuestsCanModify(true) // los invitados pueden cambiar ESE evento y nada mas
          var idcelda = sheet.getRange(last, 22, 1, 1).setValue(turnoid); // permite escribir el valor del id en una celda
        }

        else { // no se repite
          var event = calendar.createEvent(usuario, new Date(ingreso), new Date(salida), { location: sector, guests: mail, sendNotifications: true }); // Una sola vez
          var turnoid = event.getId();
          var titulo = event.setTitle(usuario + ' ' + sector);
          //var colorear = event.setColor(colorevent); // le doy color
          var descripcion = event.setDescription('Interno: ' + interno + ' Otros: ' + otrosector); // agrego una descripcion, me sirve para buscar cosas
          var modificar = event.setGuestsCanModify(true) // los invitados pueden cambiar ESE evento y nada mas
          var idcelda = sheet.getRange(last, 22, 1, 1).setValue(turnoid); // permite escribir el valor del id en una celda, actualmente la 22 es EventID

        }
        // Este es el sistema de notificaciones por mail, sólo se le notifica al encargado, puedo cambiarlo para asegurar más información
        // Esta parte me permite avisar si hay otras actividades ese día en el mismo lugar ¿No sería mejor ver si se chocan en ese horario?


        if (conflicts.length == 0) { // si no hay eventos ya
          var numeventos = 'Por ahora, no hay nadie que vaya ese día.';
        }
        else { // de otra forma quiero que los cuente
          var numeventos = 'Hay ' + conflicts.length + ' turnos en este sector reservados ese día.';
        }

        /* hasta que no lo resuelva esto no va
        if (franja.length == 0){ // si no hay asistencias en la misma franja, genero acá el mensaje a ver qué pasa
            var mismafranja = 'Ninguna persona en la misma franja horaria';
          } 
          else { // de otra forma quiero que los cuente
            var mismafranja = franja.length+' asistencias en la misma franja horaria.';
          }     
        */

        // cuerpo, asunto y data para el mail que va al responsable

        // Como agregar la URL del evento y así editarlo  
        var splitEventId = event.getId().split('@'); // saca todo lo que viene después del @
        // Open the "edit event" dialog in Calendar using this URL:
        var eventURL = "https://calendar.google.com/calendar/r/eventedit/" + Utilities.base64Encode(splitEventId[0] + " " + idcal).replace("==", ''); // Permite una URL para editar el turno

        var mensaje = usuario + ' (' + mail + ')' + ' solicita acceso a ' + sector + ' el día ' + dia + ' desde: ' + array[0][5] + ' hasta: ' + array[0][6] + frecuenciamensaje + '\n' + 'Específicamente va a: ' + interno + '. ' + numeventos + ' \n' +'\n'+ 'Editar este turno en el siguiente link o puede dirigirse directamente a su calendario para tener una mejor visión del sector: \n' + eventURL + '\n' + '\n' + 'Este es un correo automatizado, por favor no lo responda.';
        var subject = 'Solicitud de asistencia a ' + sector + ' (IFIBIO)';
        MailApp.sendEmail(referente, subject, mensaje);
      }

      else { // No está autorizado, entonces se le avisa que no va la cosa, el mensaje va al que pidió el turno, me parece fuera de lugar que vaya al responsable

        var mensaje = usuario + ' (' + mail + ') ' + ' acceso RECHAZADO a ' + sector + ' porque no está registrado este mail (' + mail + ') como USUARIO DE IFIBIO ó no fue AUTORIZADO como ingresante por parte de las autoridades. (Consulte a' + referente + ' si considera que es un error de registro o a juan.5ht@gmail.com si es un error informático) https://forms.gle/9Dw9VaHKgnFV5k8D8 \n' + '\n' + 'Este es un correo automatizado, por favor no lo responda.';
        var subject = 'IFIBIO: Solicitud de asistencia a ' + sector + ' RECHAZADO :(';
        MailApp.sendEmail(mail, subject, mensaje);
      }
      // ======== FIN SECCION DE AVISOS ==========      
      /*
      if (CONDICION1) {
      COMANDOS
      }
      else if (CONDICION2) {
      COMANDOS
      }
      else {
      COMANDOS
      }
      
      */

    } // termina el condicional de las fechas correctas
    else { // dispara el aviso al usuario que puso mal las fechas
      var mensaje = usuario + ' (' + mail + ') ' + ' acceso RECHAZADO a ' + sector + ' porque indicó mal las fechas u horarios de entrada y salida. Entrada: '+ ingreso +' y Salida: '+ salida +'\nConsidere volver a enviar el formulario. Recuerde que el primer campo es el horario de Entrada y el segundo es el de Salida y, además, que Salida es posterior a la Entrada. Se emplea un sistema de 24 horas. \n' + '\n' + 'Este es un correo automatizado, por favor no lo responda.';
      var subject = 'IFIBIO: Solicitud de asistencia a ' + sector + ' RECHAZADA por conflicto de horario :/';
      MailApp.sendEmail(mail, subject, mensaje);
    }


  }
}

function registro(e) {
  // Esta función dispara el aviso de registro en IFIBIO.
  if (e.range.getSheet().getName() == 'Datos de Usuarios') {
    // el formulario es 

    // Obtener la planilla 
    var ss = SpreadsheetApp.openById('ID_PLANILLA'); // en dónde se ejecuta el script, podría hardcodear el ID y evitar ejecuciones en cualquier lado
    // el ID de Actividades IFIBIO es: 
    var sheet = ss.getSheetByName('Datos de Usuarios'); // de dónde salen los datos, me sirve por si quiero cambiar de dónde saco los valores. Por eso le puse el nombre
    // Chequear si hay datos y obtener la última fila con valores   
    var valoresA = sheet.getRange("B1:B").getValues(); // todos los valores de la columna B, yo usaria una columna que si o si deba estar
    var last = valoresA.filter(String).length; // la ultima fila con valores de la columna B.- Hay alguna forma de hacerlo dinámico?

    /*
        Numeración de columnas
        
        A B C D E F G H I J  K  L  M  N  O  P  Q  R  S  T  U  V  W  X  Y  Z     En letras
        1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26    En números
        0 1 2 3 4 5 6 7 8 9  10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25     En arrays
                
    */
    // si agrando la cantidad de columnas a leer debo variar el último valor de la parte .getRange(last,1,1,ULTIMA_COLUMNA_PARA_LEER) 
    var data = sheet.getRange(last, 1, 1, 26).getValues(); // obtener valores (en crudo, si hay una fórmula medio que no funciona bien), encontré problemas con interpretación de fechas
    var array = sheet.getRange(last, 1, 1, 26).getDisplayValues(); // me sirve para leer contenidos que se muestran pero se calculan en el momento y es lo que vemos en la planilla

    /*  create variables 
          la variable se crea sacando la info con data[][] si viene directamente del formulario o array[][] si es resultado de otra cosa, como una fórmula
          el "0" acá representa la primera fila o columna a tomar en cuenta DENTRO del rango [FILA][COLUMNA], entonces contando desde 1 siempre es n+1
          
          si [n][n] es [0][4] = la primera fila (0+1) y la quinta columna (4+1)
    */

    var mail = data[0][1]; //mail de usuario
    var nombre = data[0][2]; //nombre del usuario


    // Este es el sistema de notificaciones por mail
    // Notifica al que autoriza J Belforte, por ahora
    var mensaje = nombre + ' (' + mail + ')' + ' solicita alta en el registro de usuarios \n' + '\n' + 'Editar la autorización en: https://docs.google.com/spreadsheets/d/12VZJEXQ5E_ioLlkFKBeXzu3Porx1Pu9gAtvDSpVrLTM/edit#gid=914328292 \n' + '\n' + 'Este es un correo automatizado, por favor no lo responda.';
    var subject = 'Registro de usuario ' + nombre;
    MailApp.sendEmail('jbelforte@fmed.uba.ar', subject, mensaje);

    // Notifica al usuario que se inscribió
    var mensajeUsr = 'Se ha registrado su solicitud de registro en IFIBIO con mail ' + mail + ' \n' + '\n' + 'Este es un correo automatizado, por favor no lo responda.';
    var subjectUsr = 'Registro de usuario IFIBIO ' + nombre;
    MailApp.sendEmail(mail, subjectUsr, mensajeUsr);

  }

}
