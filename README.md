# SCM Vidi — Package YunoHost

Application de gestion de planning et charges pour une SCM d'orthoptie.

## Installation

```bash
yunohost app install https://github.com/tanios43/scm_vidi_ynh
```

## Architecture

```
[Navigateur]  →  [Nginx]  →  [Flask/Python]  →  [SQLite]
   HTML/JS        SSOwat       API REST           data.db
```

- `/` → protégé par SSOwat, Flask injecte l'utilisateur dans le HTML
- `/api/` → bypass SSOwat, auth via token HMAC signé par Flask

## Rôles

- **Admin** (choisi à l'install) : badge ✏️, peut modifier toutes les données
- **Autres utilisateurs** : badge 👁, consultation uniquement

## Logs

```bash
journalctl -u scm_vidi -f
```

## Base de données

Stockée dans `/home/yunohost.app/scm_vidi/data.db` (SQLite).
