@app.get("/sharepoint/files")
async def sharepoint_files(folder_path: str):
    token = await get_access_token()
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{folder_path}:/children",
            headers={"Authorization": f"Bearer {token}"},
        )
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail=resp.text)
    return resp.json()


@app.get("/sharepoint/excel")
async def sharepoint_excel(item_id: str):
    token = await get_access_token()
    async with httpx.AsyncClient(follow_redirects=True) as client:
        resp = await client.get(
            f"{GRAPH_BASE}/sites/{SITE_ID}/drive/items/{item_id}/content",
            headers={"Authorization": f"Bearer {token}"},
        )
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail=resp.text)

    df = pd.read_excel(io.BytesIO(resp.content))
    return df.to_dict(orient="records")
