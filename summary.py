# summary.py
import re
from collections import defaultdict

ENV_NORMALIZE = {
    'prod': 'Production',
    'production': 'Production',
    'test': 'Test',
    'uat': 'UAT',
    'dev': 'Dev',
}

DB_NAME_PATTERN = re.compile(r'(^|[^a-z])(db|sql|ora)(\d+)?($|[^a-z])', re.IGNORECASE)

def normalize_env(env: str) -> str:
    if not env:
        return 'Unknown'
    key = env.strip().lower()
    return ENV_NORMALIZE.get(key, env.strip().title())

def is_db_host(server_name: str) -> bool:
    if not server_name:
        return False
    return bool(DB_NAME_PATTERN.search(server_name))

def summarize(rows):
    """
    rows: iterable[dict] with keys:
      ESATS_ID, APPLICATION, SERVER, ENVIRONMENT, OS_NAME, OS_VERSION
    """
    envs = set()
    servers_by_env = defaultdict(set)
    os_names = set()
    os_versions = set()
    db_by_env = defaultdict(set)

    for r in rows:
        env = normalize_env(r.get('ENVIRONMENT', ''))
        server = (r.get('SERVER') or '').strip()
        os_name = (r.get('OS_NAME') or '').strip()
        os_ver = (r.get('OS_VERSION') or '').strip()

        if env != 'Unknown':
            envs.add(env)

        if server:
            servers_by_env[env].add(server)
            if is_db_host(server):
                db_by_env[env].add(server)

        if os_name:
            os_names.add(os_name)
        if os_ver:
            os_versions.add(os_ver)

    return {
        "Environments": sorted(envs),
        "Servers by Environment": {
            env: sorted(list(servers))
            for env, servers in sorted(servers_by_env.items())
            if servers
        },
        "Operating System": sorted(os_names),
        "OS Version": sorted(os_versions),
        "Database Servers by Environment": {
            env: sorted(list(db_hosts))
            for env, db_hosts in sorted(db_by_env.items())
            if db_hosts
        }
    }
