# excel-ai-gpt-addin

## Caching des fonctions IA

- Le cache est désormais persistant par défaut (TTL 24h) : les résultats sont réutilisés après un rechargement d'Excel sans refaire un appel IA.
- Vous pouvez forcer le mode "cache only" via les options JSON des fonctions (ex.: `{ "cacheOnly": true }`). Si aucune valeur n'est trouvée en cache, la fonction renvoie `#AI_CACHE_MISS` au lieu de déclencher un nouvel appel. Recalculez ensuite manuellement (F9 ou modifier les options) si vous voulez vraiment relancer la requête IA.