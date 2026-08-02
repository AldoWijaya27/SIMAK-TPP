import json
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Protocol
from urllib import error, request


class UserAccessGateway(Protocol):
    def is_allowed(self, username: str) -> bool | None:
        ...


@dataclass
class FirebaseRealtimeUserAccessGateway:
    base_url: str
    timeout_seconds: int = 8

    def is_allowed(self, username: str) -> bool | None:
        """
        Return:
          True  → akses diizinkan
          False → akses ditolak (username tidak terdaftar)
          None  → tidak bisa konek ke server (no internet / timeout / firewall)
        """
        username = (username or "").strip().lower()
        if not username:
            return False

        url = f"{self.base_url.rstrip('/')}/allowed_users/{username}.json"
        req = request.Request(url=url, method="GET")
        import ssl
        try:
            context = ssl._create_unverified_context()
            with request.urlopen(req, timeout=self.timeout_seconds, context=context) as response:
                body = response.read().decode("utf-8", errors="ignore").strip()
        except (error.URLError, error.HTTPError, TimeoutError, ValueError):
            # Tidak bisa konek ke Firebase — kembalikan None agar pemanggil bisa
            # menampilkan pesan error yang spesifik (bukan sekadar "ERROR!")
            return None

        if not body:
            return False

        try:
            data = json.loads(body)
        except json.JSONDecodeError:
            return False

        return data is True


@dataclass
class ValidateUserAccess:
    gateway: UserAccessGateway

    def execute(self, username: str) -> bool | None:
        # timestamp disiapkan untuk extensibility audit ke depan.
        _ = datetime.now(timezone.utc).isoformat()
        return self.gateway.is_allowed(username)
