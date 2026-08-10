use serde_json::{json, Value};
use std::env;
use std::io::{Read, Write};
use std::net::{Ipv4Addr, SocketAddrV4, TcpStream};
use std::time::{Duration, Instant};

const PORT_ENV: &str = "OOXML_ACTIVE_CONTEXT_BRIDGE_PORT";
const TOKEN_ENV: &str = "OOXML_ACTIVE_CONTEXT_BRIDGE_TOKEN";
const MAX_RESPONSE_BYTES: usize = 8 * 1_024 * 1_024;
const IO_TIMEOUT: Duration = Duration::from_secs(2);

pub struct SelectionTools;

impl SelectionTools {
    pub fn get_active_context() -> String {
        let port = match env::var(PORT_ENV) {
            Ok(value) => match value.parse::<u16>() {
                Ok(port) if port > 0 => port,
                _ => return unavailable(
                    "active_context_bridge_unavailable",
                    "The active Viewer context bridge is not available. Launch this MCP server through the OOXML Viewer VS Code extension.",
                ),
            },
            Err(_) => return unavailable(
                "active_context_bridge_unavailable",
                "The active Viewer context bridge is not available. Launch this MCP server through the OOXML Viewer VS Code extension.",
            ),
        };
        let token = match env::var(TOKEN_ENV) {
            Ok(value) if valid_token(&value) => value,
            _ => return unavailable(
                "active_context_bridge_unavailable",
                "The active Viewer context bridge is not available. Launch this MCP server through the OOXML Viewer VS Code extension.",
            ),
        };

        match fetch_active_context(port, &token) {
            Ok(value) => value.to_string(),
            Err(message) => unavailable("active_context_bridge_error", &message),
        }
    }
}

fn valid_token(token: &str) -> bool {
    token.len() == 64 && token.bytes().all(|byte| byte.is_ascii_hexdigit())
}

fn unavailable(reason: &str, message: &str) -> String {
    json!({
        "schemaVersion": 1,
        "available": false,
        "context": Value::Null,
        "reason": reason,
        "message": message,
    })
    .to_string()
}

fn fetch_active_context(port: u16, token: &str) -> Result<Value, String> {
    fetch_active_context_with_timeout(port, token, IO_TIMEOUT)
}

fn remaining(deadline: Instant) -> Result<Duration, String> {
    deadline
        .checked_duration_since(Instant::now())
        .filter(|duration| !duration.is_zero())
        .ok_or_else(|| "The selection bridge request timed out.".to_string())
}

fn fetch_active_context_with_timeout(
    port: u16,
    token: &str,
    timeout: Duration,
) -> Result<Value, String> {
    if port == 0 || !valid_token(token) {
        return Err("Selection bridge credentials are invalid.".to_string());
    }
    let deadline = Instant::now() + timeout;
    let address = SocketAddrV4::new(Ipv4Addr::LOCALHOST, port);
    let mut stream = TcpStream::connect_timeout(&address.into(), remaining(deadline)?)
        .map_err(|_| "Cannot connect to the VS Code selection bridge.".to_string())?;
    stream
        .set_read_timeout(Some(remaining(deadline)?))
        .map_err(|_| "Cannot configure the selection bridge connection.".to_string())?;
    stream
        .set_write_timeout(Some(remaining(deadline)?))
        .map_err(|_| "Cannot configure the selection bridge connection.".to_string())?;

    let request = format!(
        "GET /v1/context HTTP/1.1\r\nHost: 127.0.0.1:{port}\r\nAuthorization: Bearer {token}\r\nAccept: application/json\r\nConnection: close\r\n\r\n"
    );
    stream
        .write_all(request.as_bytes())
        .map_err(|_| "Cannot request the active selection.".to_string())?;

    let mut response = Vec::new();
    let mut buffer = [0u8; 8 * 1_024];
    loop {
        stream
            .set_read_timeout(Some(remaining(deadline)?))
            .map_err(|_| "Cannot configure the selection bridge connection.".to_string())?;
        let count = stream
            .read(&mut buffer)
            .map_err(|_| "Cannot read the active selection response.".to_string())?;
        if count == 0 {
            break;
        }
        response.extend_from_slice(&buffer[..count]);
        if response.len() > MAX_RESPONSE_BYTES {
            return Err("The active selection response exceeded the resource limit.".to_string());
        }
    }
    if response.len() > MAX_RESPONSE_BYTES {
        return Err("The active selection response exceeded the resource limit.".to_string());
    }
    parse_response(&response)
}

fn parse_response(response: &[u8]) -> Result<Value, String> {
    let separator = response
        .windows(4)
        .position(|window| window == b"\r\n\r\n")
        .ok_or_else(|| "The selection bridge returned an invalid HTTP response.".to_string())?;
    let header = std::str::from_utf8(&response[..separator])
        .map_err(|_| "The selection bridge returned invalid HTTP headers.".to_string())?;
    let mut lines = header.split("\r\n");
    let status = lines.next().unwrap_or_default();
    if status != "HTTP/1.1 200 OK" && status != "HTTP/1.0 200 OK" {
        return Err("The selection bridge rejected the request.".to_string());
    }

    let mut content_length = None;
    for line in lines {
        let Some((name, value)) = line.split_once(':') else {
            return Err("The selection bridge returned invalid HTTP headers.".to_string());
        };
        if name.eq_ignore_ascii_case("transfer-encoding") {
            return Err("The selection bridge returned an unsupported HTTP response.".to_string());
        }
        if name.eq_ignore_ascii_case("content-length") {
            let length = value.trim().parse::<usize>().map_err(|_| {
                "The selection bridge returned an invalid content length.".to_string()
            })?;
            content_length = Some(length);
        }
    }
    let body = &response[separator + 4..];
    let expected = content_length
        .ok_or_else(|| "The selection bridge omitted the content length.".to_string())?;
    if expected != body.len() || expected > MAX_RESPONSE_BYTES {
        return Err("The selection bridge returned an incomplete response.".to_string());
    }

    let value: Value = serde_json::from_slice(body)
        .map_err(|_| "The selection bridge returned invalid JSON.".to_string())?;
    if value.get("schemaVersion").and_then(Value::as_u64) != Some(1)
        || value.get("available").and_then(Value::as_bool) != Some(true)
        || !value
            .as_object()
            .is_some_and(|object| object.contains_key("context"))
    {
        return Err("The selection bridge returned an unsupported payload.".to_string());
    }
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::net::TcpListener;
    use std::thread;

    const TOKEN: &str = "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef";

    fn serve_once(status: &str, body: &str) -> (u16, thread::JoinHandle<()>) {
        let listener = TcpListener::bind((Ipv4Addr::LOCALHOST, 0)).unwrap();
        let port = listener.local_addr().unwrap().port();
        let status = status.to_string();
        let body = body.to_string();
        let handle = thread::spawn(move || {
            let (mut stream, _) = listener.accept().unwrap();
            stream.set_read_timeout(Some(IO_TIMEOUT)).unwrap();
            let mut request = [0u8; 1_024];
            let count = stream.read(&mut request).unwrap();
            let request = std::str::from_utf8(&request[..count]).unwrap();
            assert!(request.starts_with("GET /v1/context HTTP/1.1\r\n"));
            assert!(request.contains(&format!("Authorization: Bearer {TOKEN}\r\n")));
            let response = format!(
                "HTTP/1.1 {status}\r\nContent-Type: application/json\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{body}",
                body.len(),
            );
            stream.write_all(response.as_bytes()).unwrap();
        });
        (port, handle)
    }

    #[test]
    fn reads_a_valid_bounded_active_context_payload() {
        let body = r#"{"schemaVersion":1,"available":true,"context":{"document":{"format":"docx","name":"a.docx","path":"/tmp/a.docx"},"view":{"format":"docx","pageIndex":2},"selection":{"format":"docx","kind":"text","text":"hello"}}}"#;
        let (port, server) = serve_once("200 OK", body);
        let value = fetch_active_context(port, TOKEN).unwrap();
        server.join().unwrap();
        assert_eq!(value["context"]["selection"]["text"], "hello");
    }

    #[test]
    fn preserves_an_explicit_no_active_preview_result() {
        let body = r#"{"schemaVersion":1,"available":true,"context":null}"#;
        let (port, server) = serve_once("200 OK", body);
        let value = fetch_active_context(port, TOKEN).unwrap();
        server.join().unwrap();
        assert!(value["context"].is_null());
    }

    #[test]
    fn rejects_http_errors_and_unsupported_payloads() {
        let (port, server) = serve_once("401 Unauthorized", r#"{"error":"unauthorized"}"#);
        assert!(fetch_active_context(port, TOKEN).is_err());
        server.join().unwrap();

        let (port, server) = serve_once(
            "200 OK",
            r#"{"schemaVersion":2,"available":true,"context":null}"#,
        );
        assert!(fetch_active_context(port, TOKEN).is_err());
        server.join().unwrap();
    }

    #[test]
    fn rejects_request_injection_in_the_token() {
        assert!(!valid_token(&format!("{}\r\nInjected: yes", TOKEN)));
        assert!(fetch_active_context(1, "not-a-token").is_err());
    }

    #[test]
    fn enforces_a_total_deadline_against_a_slow_trickle_response() {
        let listener = TcpListener::bind((Ipv4Addr::LOCALHOST, 0)).unwrap();
        let port = listener.local_addr().unwrap().port();
        let server = thread::spawn(move || {
            let (mut stream, _) = listener.accept().unwrap();
            let mut request = [0u8; 1_024];
            let _ = stream.read(&mut request);
            for byte in b"HTTP/1.1 200 OK\r\nContent-Length: 2\r\n\r\n{}" {
                if stream.write_all(&[*byte]).is_err() {
                    break;
                }
                thread::sleep(Duration::from_millis(20));
            }
        });

        let started = Instant::now();
        let result = fetch_active_context_with_timeout(port, TOKEN, Duration::from_millis(100));
        assert!(result.is_err());
        assert!(started.elapsed() < Duration::from_secs(1));
        server.join().unwrap();
    }
}
