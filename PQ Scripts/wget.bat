del LPQ.pq
del Inno.Online.pq

curl -L "https://api.github.com/gists/700a6d65e098189881ecd77e585b233a" -H "Content-Type:application/json" | jq -rc .files[\"LPQ.pq\"].content >> LPQ.pq

curl -L "https://api.github.com/gists/700a6d65e098189881ecd77e585b233a" -H "Content-Type:application/json" | jq -rc .files[\"Inno.Online.pq\"].content >> Inno.Online.pq
