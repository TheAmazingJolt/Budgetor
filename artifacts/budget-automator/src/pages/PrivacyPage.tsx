import { DollarSign } from "lucide-react";

export function PrivacyPage() {
  return (
    <div className="min-h-screen bg-white text-slate-800">
      <header className="border-b border-slate-200 px-6 py-4 flex items-center gap-3">
        <a href="/" className="flex items-center gap-2 text-slate-700 hover:text-slate-900">
          <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-primary to-emerald-600 flex items-center justify-center">
            <DollarSign className="w-4 h-4 text-white" />
          </div>
          <span className="font-semibold text-lg">Budgify</span>
        </a>
      </header>

      <main className="max-w-3xl mx-auto px-6 py-12">
        <h1 className="text-3xl font-bold mb-2">Privacy Policy</h1>
        <p className="text-slate-500 mb-8">Last updated: April 23, 2026</p>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Overview</h2>
          <p className="text-slate-600 leading-relaxed">
            Budgify ("we", "us", "our") is a personal budgeting tool. This Privacy Policy explains what
            information we collect, how we use it, and your rights. We do not sell your personal data.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Information We Collect</h2>
          <ul className="list-disc pl-6 space-y-2 text-slate-600 leading-relaxed">
            <li><strong>Account information:</strong> Your email address and display name when you create an account or sign in via Google or Apple.</li>
            <li><strong>Budget data:</strong> Bills, debts, income sources, savings goals, weekly budgets, and check-in records that you enter into the app. This data is stored in our database to power the service.</li>
            <li><strong>Google account access:</strong> If you connect Google Sheets sync, we receive an OAuth token scoped to your Google Sheets. We use this only to read and write the specific spreadsheet you authorize. We do not access any other Google data.</li>
            <li><strong>Microsoft account access:</strong> If you connect OneDrive sync, we receive an OAuth token scoped to your OneDrive files. We use this only to create and update the Excel file you authorize. We do not access any other Microsoft data.</li>
            <li><strong>Payment information:</strong> Subscription billing is handled entirely by Stripe. We do not store your credit card number or payment details. We receive only a customer ID and subscription status from Stripe.</li>
            <li><strong>Usage data:</strong> Basic server logs (IP address, timestamps, request paths) are retained for security and debugging. We do not use third-party analytics trackers.</li>
          </ul>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">How We Use Your Information</h2>
          <ul className="list-disc pl-6 space-y-2 text-slate-600 leading-relaxed">
            <li>To operate and maintain your Budgify account and provide the budgeting service.</li>
            <li>To sync your budget to Google Sheets or OneDrive when you enable those integrations.</li>
            <li>To process payments and manage your subscription via Stripe.</li>
            <li>To send transactional emails such as password reset links and billing receipts.</li>
            <li>To investigate and prevent fraud, abuse, or security incidents.</li>
          </ul>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Third-Party Services</h2>
          <p className="text-slate-600 leading-relaxed mb-3">
            We rely on the following third parties to operate the service:
          </p>
          <ul className="list-disc pl-6 space-y-2 text-slate-600 leading-relaxed">
            <li><strong>Google LLC</strong> — OAuth authentication and Google Sheets API. Subject to Google's Privacy Policy.</li>
            <li><strong>Microsoft Corporation</strong> — OAuth authentication and Microsoft Graph API for OneDrive. Subject to Microsoft's Privacy Policy.</li>
            <li><strong>Stripe, Inc.</strong> — Payment processing and subscription management. Subject to Stripe's Privacy Policy.</li>
            <li><strong>Railway Corp.</strong> — Application and database hosting. Data is stored in the United States.</li>
          </ul>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Data Retention</h2>
          <p className="text-slate-600 leading-relaxed">
            We retain your account and budget data for as long as your account is active. If you delete your
            account, your personal data and budget records are permanently deleted within 30 days.
            Anonymized aggregate statistics (e.g. total number of users) may be retained indefinitely.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Your Rights</h2>
          <ul className="list-disc pl-6 space-y-2 text-slate-600 leading-relaxed">
            <li><strong>Access:</strong> You can view and export your budget data at any time from the Download section of the app.</li>
            <li><strong>Correction of budget data:</strong> You can edit your bills, debts, income, savings goals, and weekly budgets directly in the app at any time.</li>
            <li><strong>Correction of account email:</strong> To change the email address on your account, email us at{" "}
              <a href="mailto:support@budgify.org" className="text-primary underline">support@budgify.org</a>{" "}
              from your existing account email. We will process the change within 30 days.</li>
            <li><strong>Deletion:</strong> To delete your account and all associated data, email us at{" "}
              <a href="mailto:support@budgify.org" className="text-primary underline">support@budgify.org</a>.
              We will permanently remove your account and all budget, debt, savings, and sync-integration data
              within 30 days of verification.</li>
            <li><strong>Revoke integrations:</strong> You can disconnect Google or Microsoft access at any time from your Google/Microsoft account security settings, or by disconnecting the integration from within the app.</li>
          </ul>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Children's Privacy</h2>
          <p className="text-slate-600 leading-relaxed">
            Budgify is not directed to children under 13. We do not knowingly collect personal information
            from children. If you believe a child has created an account, please contact us and we will
            promptly delete it.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Changes to This Policy</h2>
          <p className="text-slate-600 leading-relaxed">
            We may update this policy from time to time. When we make material changes, we will notify
            you by email or by displaying a notice in the app. Continued use of Budgify after the effective
            date of any change constitutes acceptance of the updated policy.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Contact</h2>
          <p className="text-slate-600 leading-relaxed">
            Questions or requests regarding your privacy can be sent to{" "}
            <a href="mailto:support@budgify.org" className="text-primary underline">support@budgify.org</a>.
          </p>
        </section>
      </main>

      <footer className="border-t border-slate-200 px-6 py-6 text-center text-sm text-slate-400">
        <div className="flex justify-center gap-6 mb-2">
          <a href="/" className="hover:text-slate-600">Home</a>
          <a href="/privacy" className="hover:text-slate-600">Privacy Policy</a>
          <a href="/terms" className="hover:text-slate-600">Terms of Service</a>
        </div>
        © {new Date().getFullYear()} Budgify. All rights reserved.
      </footer>
    </div>
  );
}
