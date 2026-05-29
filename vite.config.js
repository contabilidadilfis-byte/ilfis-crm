# 🔥 ILFIS CRM — Guía de Instalación y Despliegue

CRM comercial con sincronización en tiempo real vía Firebase Firestore.
Todos los asesores ven y editan los mismos datos desde cualquier dispositivo.

---

## ⏱ Tiempo estimado: 20–30 minutos (sin experiencia técnica previa)

---

## PASO 1 — Crear la base de datos en Firebase (gratis)

1. Ve a **https://console.firebase.google.com**
2. Clic en **"Agregar proyecto"**
3. Nombre del proyecto: `ilfis-crm` → **Continuar**
4. Desactiva Google Analytics → **Crear proyecto**
5. Espera que se cree y clic en **Continuar**

### Crear la base de datos Firestore:
6. En el menú izquierdo → **"Firestore Database"**
7. Clic en **"Crear base de datos"**
8. Selecciona **"Iniciar en modo de prueba"** → **Siguiente**
9. Elige la región más cercana (ej: `nam5 (us-central)`) → **Listo**

### Obtener las credenciales:
10. En el menú izquierdo → ⚙️ **Configuración del proyecto**
11. Baja a **"Tus apps"** → clic en el ícono **`</>`** (Web)
12. Nombre de la app: `ILFIS CRM` → **Registrar app**
13. Verás un bloque `firebaseConfig` con tus credenciales → **cópialo**

---

## PASO 2 — Configurar el proyecto

1. Abre la carpeta `ilfis-crm` en tu computador
2. Copia el archivo `.env.example` y renómbralo a `.env`
3. Abre `.env` con el Bloc de notas y reemplaza los valores con los de tu `firebaseConfig`:

```
VITE_FIREBASE_API_KEY=      ← valor de "apiKey"
VITE_FIREBASE_AUTH_DOMAIN=  ← valor de "authDomain"
VITE_FIREBASE_PROJECT_ID=   ← valor de "projectId"
VITE_FIREBASE_STORAGE_BUCKET= ← valor de "storageBucket"
VITE_FIREBASE_MESSAGING_SENDER_ID= ← valor de "messagingSenderId"
VITE_FIREBASE_APP_ID=       ← valor de "appId"
```

---

## PASO 3 — Instalar y ejecutar localmente (opcional)

Necesitas tener **Node.js** instalado (https://nodejs.org — versión LTS).

```bash
# Dentro de la carpeta ilfis-crm:
npm install
npm run dev
```

Abre **http://localhost:5173** en tu navegador.
La primera vez que cargue, los datos de ejemplo se crearán automáticamente en Firestore.

---

## PASO 4 — Publicar en internet con Vercel (gratis)

Los asesores acceden con un link desde cualquier navegador, sin instalar nada.

1. Crea una cuenta gratis en **https://vercel.com** (puedes entrar con GitHub)
2. Sube la carpeta `ilfis-crm` a un repositorio en **https://github.com**
   - Crea cuenta en GitHub si no tienes
   - Crea repositorio nuevo → sube todos los archivos
3. En Vercel → **"Add New Project"** → importa el repositorio de GitHub
4. En la sección **"Environment Variables"**, agrega las mismas variables de tu `.env`
5. Clic en **"Deploy"**

✅ En 2 minutos tendrás un link tipo `ilfis-crm.vercel.app` para compartir con los asesores.

---

## PASO 5 — Compartir con los asesores

Simplemente envía el link a cada asesor. No necesitan instalar nada.

En la pestaña **"👥 Asesores"** del CRM puedes:
- Agregar o eliminar asesores
- Cambiar quién está en sesión (selector "Sesión activa como")
- Ver el rendimiento de cada asesor en tiempo real

---

## 🔒 Seguridad (recomendado antes de uso en producción)

El modo de prueba de Firestore expira en 30 días. Para renovarlo o protegerlo:

1. Ve a Firestore → **"Reglas"**
2. Reemplaza las reglas con:

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /{document=**} {
      allow read, write: if true; // Cambiar por autenticación en producción
    }
  }
}
```

Para un CRM de uso interno sin datos sensibles, el modo de prueba es suficiente.

---

## 📊 Estructura de la base de datos

Firestore crea automáticamente 3 colecciones:

| Colección   | Descripción                              |
|-------------|------------------------------------------|
| `clients`   | Todos los clientes y oportunidades       |
| `payments`  | Registros de pago y datos de facturación |
| `advisors`  | Lista de asesores del equipo             |

---

## ❓ Soporte

Si algo no funciona, revisa:
- Que el archivo `.env` tenga los valores correctos (sin comillas, sin espacios)
- Que Firestore esté en modo "prueba" (no "producción bloqueada")
- La consola del navegador (F12) para ver errores específicos
