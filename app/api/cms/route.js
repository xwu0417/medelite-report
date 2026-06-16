export async function GET(request) {
  const { searchParams } = new URL(request.url);
  const endpoint = searchParams.get("endpoint");
  
  const res = await fetch(endpoint, {
    headers: { "Content-Type": "application/json" },
  });
  
  const data = await res.json();
  return Response.json(data);
}