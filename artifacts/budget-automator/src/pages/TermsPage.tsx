import { DollarSign } from "lucide-react";

export function TermsPage() {
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
        <h1 className="text-3xl font-bold mb-2">Terms of Service</h1>
        <p className="text-slate-500 mb-8">Last updated: April 23, 2026</p>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Acceptance of Terms</h2>
          <p className="text-slate-600 leading-relaxed">
            By creating an account or using Budgify ("the Service"), you agree to these Terms of Service.
            If you do not agree, do not use the Service. These terms constitute a binding agreement between
            you and Budgify.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">The Service</h2>
          <p className="text-slate-600 leading-relaxed">
            Budgify is a personal budgeting application that helps you create weekly budgets, track bills
            and debts, set savings goals, and optionally sync your budget to Google Sheets or Microsoft
            OneDrive. The Service is provided on a free and paid (Pro) tier basis.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Accounts</h2>
          <ul className="list-disc pl-6 space-y-2 text-slate-600 leading-relaxed">
            <li>You must provide accurate and current information when creating an account.</li>
            <li>You are responsible for maintaining the security of your account credentials.</li>
            <li>You may not share your account with others or create accounts on behalf of other people without their consent.</li>
            <li>Guest accounts are temporary; data is not guaranteed to be preserved without creating a permanent account.</li>
          </ul>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Subscriptions and Billing</h2>
          <ul className="list-disc pl-6 space-y-2 text-slate-600 leading-relaxed">
            <li>The free tier provides access to core budgeting features with limits on the number of bills, debts, and savings goals.</li>
            <li>The Pro tier is a paid monthly or annual subscription that removes those limits and unlocks additional features.</li>
            <li>Subscriptions are billed through Stripe. By subscribing, you authorize recurring charges to your payment method.</li>
            <li>You may cancel your subscription at any time from your account settings. Cancellation takes effect at the end of the current billing period; no partial refunds are provided.</li>
            <li>We reserve the right to change pricing with at least 30 days' notice to existing subscribers.</li>
          </ul>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Acceptable Use</h2>
          <p className="text-slate-600 leading-relaxed mb-3">You agree not to:</p>
          <ul className="list-disc pl-6 space-y-2 text-slate-600 leading-relaxed">
            <li>Use the Service for any unlawful purpose or in violation of any applicable law.</li>
            <li>Attempt to reverse engineer, decompile, or extract the source code of the Service.</li>
            <li>Abuse, harass, or harm other users of the Service.</li>
            <li>Use automated scripts or bots to access the Service in a way that degrades performance for others.</li>
            <li>Circumvent any access controls or security measures.</li>
          </ul>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Third-Party Integrations</h2>
          <p className="text-slate-600 leading-relaxed">
            The Service integrates with Google Sheets and Microsoft OneDrive. Your use of those integrations
            is also governed by Google's and Microsoft's respective terms of service. We are not responsible
            for the availability, accuracy, or conduct of those third-party services.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Your Data</h2>
          <p className="text-slate-600 leading-relaxed">
            You retain ownership of all budget data you enter into the Service. By using the Service, you
            grant us a limited license to store and process that data solely for the purpose of operating the
            Service on your behalf. We do not claim any intellectual property rights over your data.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Disclaimer of Warranties</h2>
          <p className="text-slate-600 leading-relaxed">
            The Service is provided "as is" and "as available" without any warranty of any kind, express or
            implied, including but not limited to warranties of merchantability, fitness for a particular
            purpose, or non-infringement. We do not warrant that the Service will be uninterrupted, error-free,
            or that any defects will be corrected. Financial decisions are your own responsibility; Budgify is
            a tool only and does not constitute financial advice.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Limitation of Liability</h2>
          <p className="text-slate-600 leading-relaxed">
            To the maximum extent permitted by law, Budgify shall not be liable for any indirect, incidental,
            special, consequential, or punitive damages arising out of or related to your use of the Service,
            including but not limited to loss of data, loss of profits, or financial loss. Our total liability
            to you for any claim shall not exceed the amount you paid us in the 12 months preceding the claim.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Account Termination</h2>
          <p className="text-slate-600 leading-relaxed">
            You may delete your account at any time from your account settings. We may suspend or terminate
            your account if you violate these Terms. Upon termination, your data will be deleted in accordance
            with our Privacy Policy.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Changes to These Terms</h2>
          <p className="text-slate-600 leading-relaxed">
            We may update these Terms from time to time. We will notify you of material changes via email or
            an in-app notice. Continued use of the Service after the effective date of any change constitutes
            acceptance of the updated Terms.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Governing Law</h2>
          <p className="text-slate-600 leading-relaxed">
            These Terms are governed by the laws of the United States. Any disputes arising under these Terms
            shall be resolved through binding arbitration or in the courts of competent jurisdiction in the
            United States, at our election.
          </p>
        </section>

        <section className="mb-8">
          <h2 className="text-xl font-semibold mb-3">Contact</h2>
          <p className="text-slate-600 leading-relaxed">
            Questions about these Terms can be sent to{" "}
            <a href="mailto:legal@budgify.org" className="text-primary underline">legal@budgify.org</a>.
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
