import MailboxForm from "@/components/mailbox-form";
import MicrosoftMailbox from "@/components/microsoft-mailbox";
import { PageHeader } from "@/components/ui";
import { PROVIDERS } from "@/lib/providers";

export default function NewMailboxPage() {
  return (
    <div className="mx-auto max-w-xl space-y-6">
      <PageHeader
        title="Add mailbox"
        subtitle="Connect over IMAP or POP3. The login is checked first, then saved encrypted and scanned right away."
      />
      {process.env.MICROSOFT_CLIENT_ID ? <MicrosoftMailbox /> : null}
      <MailboxForm providers={PROVIDERS} />
    </div>
  );
}
