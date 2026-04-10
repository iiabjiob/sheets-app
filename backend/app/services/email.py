from __future__ import annotations

import smtplib
from email.message import EmailMessage

from app.core.config import get_settings
from app.core.logger import get_logger

logger = get_logger("mail")


class EmailService:
    def send_registration_verification(
        self,
        *,
        email: str,
        full_name: str,
        verification_url: str,
    ) -> None:
        settings = get_settings()
        subject = "Confirm your registration"
        body = (
            f"Hello {full_name},\n\n"
            "Confirm your email to activate your account.\n\n"
            f"Verification link: {verification_url}\n\n"
            "If you did not request this registration, ignore this message."
        )
        delivery_mode = settings.email_delivery_mode.strip().lower()

        if delivery_mode == "log":
            self._log_email(subject=subject, email=email, body=body)
            return

        if delivery_mode not in {"auto", "smtp"}:
            raise RuntimeError(
                "EMAIL_DELIVERY_MODE must be one of: log, auto, smtp. "
                f"Got: {settings.email_delivery_mode!r}"
            )

        if not settings.smtp_host or not settings.smtp_from_email:
            if delivery_mode == "smtp":
                raise RuntimeError(
                    "SMTP delivery is enabled, but SMTP_HOST / SMTP_FROM_EMAIL are not configured."
                )

            self._log_email(subject=subject, email=email, body=body)
            return

        message = EmailMessage()
        message["Subject"] = subject
        message["From"] = settings.smtp_from_email
        message["To"] = email
        message.set_content(body)

        if settings.smtp_use_ssl:
            with smtplib.SMTP_SSL(settings.smtp_host, settings.smtp_port) as client:
                self._login_if_needed(client)
                client.send_message(message)
            return

        with smtplib.SMTP(settings.smtp_host, settings.smtp_port) as client:
            if settings.smtp_use_tls:
                client.starttls()
            self._login_if_needed(client)
            client.send_message(message)

    def _log_email(self, *, subject: str, email: str, body: str) -> None:
        logger.info(
            "Email delivery mode=log. Registration email to %s\nSubject: %s\n\n%s",
            email,
            subject,
            body,
        )

    def _login_if_needed(self, client: smtplib.SMTP) -> None:
        settings = get_settings()
        if settings.smtp_username and settings.smtp_password:
            client.login(settings.smtp_username, settings.smtp_password)


email_service = EmailService()