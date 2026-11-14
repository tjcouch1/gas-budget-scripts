/**
 * Receipt email fixtures: the single source of sample emails for the receipt-parsing unit tests,
 * the other tests that need a receipt thread, and the local demo backend's fake Gmail inbox.
 *
 * If a real email format changes, change (or add) the fixture here. Each fixture keeps the email
 * and the parsed result the code is expected to produce next to each other. Pure data: no mocks
 * or GAS globals, so anything (unit tests, demo server, e2e) can import it.
 *
 * Names and addresses are obviously fake. The vendor sender addresses are the public ones the
 * parsers match on.
 */

export const CHASE = "Chase <no.reply.alerts@chase.com>";
export const PAYPAL = "PayPal <service@paypal.com>";
/** PayPal sometimes puts the address itself, quoted, in the display name */
export const PAYPAL_QUOTED = '"service@paypal.com" <service@paypal.com>';
export const VENMO = "Venmo <venmo@venmo.com>";
/** The person who forwards receipts (ForwardEmailAddress / ForwardName in the test variables) */
export const FWD = "Sam Example <fwd@example.com>";

/** A Gmail message to build a thread from (structurally the mock's `MessageInit`) */
export interface EmailFixture {
  id?: string;
  from: string;
  to?: string;
  subject: string;
  body: string;
  date?: Date;
}

/** What the parser should produce. Only the listed keys are asserted */
export interface ExpectedReceipt {
  cost?: number;
  name?: string;
  type?: string;
  category?: string;
  date?: Date;
  /** Asserts `cost` is NaN (a `$abc` cost that parseFloat cannot read) */
  costIsNaN?: true;
}

export interface ReceiptFixture {
  /** Used as the unit test's name */
  description: string;
  email: EmailFixture;
  /** `null`: not a receipt, so the thread is dropped (it only gets a note) */
  expected: ExpectedReceipt | null;
}

/** Fixed email date for the parsing tests */
export const FIXTURE_DATE = new Date("2024-03-01T18:30:00Z");

const PAYPAL_PAYMENT_BODY =
  "Hello,\n\nYou sent $25.50 to Jane Doe\n\nTransaction ID\nABC123\n\nLunch money\nThanks\n";

const FORWARDED_PAYPAL_EXPECTED = {
  cost: 25.5,
  name: "Jane Doe - Lunch money",
};

/** A reply from someone with no rules, in a thread the query pulled in */
const UNKNOWN_SENDER_REPLY: EmailFixture = {
  from: "Pat Friend <pat.friend@example.net>",
  subject: "Re: Dinner Friday?",
  body: "Count me in!\n",
};

/** The forwarder's own reply (no forward prefix) */
const FORWARDER_REPLY: EmailFixture = {
  from: FWD,
  subject: "Re: Dinner Friday?",
  body: "Sounds good, see you there.\n",
};

export const receiptFixtures = {
  chasePurchase: {
    description: "parses a normal purchase from the subject",
    email: {
      from: CHASE,
      subject: "You made a $12.34 transaction with Trader Joe's",
      body: "irrelevant",
      date: FIXTURE_DATE,
    },
    expected: {
      cost: 12.34,
      name: "Trader Joe's",
      type: "Credit",
      category: undefined,
      date: FIXTURE_DATE,
    },
  },
  chasePendingCredit: {
    description:
      "parses a pending credit as a negative cost with merchant from the body",
    email: {
      from: CHASE,
      subject: "You have a $45.00 credit pending on your credit card",
      body: "Account ending in 1234\r\nMerchant   Target Store  \r\nDate Mar 1\r\n",
    },
    expected: { cost: -45, name: "Target Store", type: "Credit" },
  },
  chaseGasStation: {
    description: "parses a gas station receipt with the Gas category",
    email: {
      from: CHASE,
      subject: "You used your card at a gas station",
      body: "Hello\nMerchant   SHELL OIL 123  \n\nMore text\n",
    },
    expected: {
      cost: undefined,
      name: "SHELL OIL 123 (Gas)",
      category: "Gas",
      type: "Credit",
    },
  },
  chaseStatement: {
    description:
      "treats an unrecognized Chase subject as a non-receipt with a note",
    email: {
      from: CHASE,
      subject: "Your statement is ready",
      body: "x".repeat(500),
      date: FIXTURE_DATE,
    },
    expected: null,
  },

  paypalPaymentSent: {
    description: "parses a payment sent",
    email: {
      from: PAYPAL,
      to: "someone@example.com",
      subject: "You sent a $25.50 USD payment",
      body: PAYPAL_PAYMENT_BODY,
    },
    // Default prefix is TJ when 'to' is not the forward address
    expected: { cost: 25.5, name: "Jane Doe - Lunch money", type: "TJ Paypal" },
  },
  paypalPaymentSentToForwarder: {
    description:
      "uses ForwardName as type prefix when sent to the forward address",
    email: {
      from: PAYPAL,
      to: "Sam <fwd@example.com>",
      subject: "You sent a $25.50 USD payment",
      body: PAYPAL_PAYMENT_BODY,
    },
    expected: { type: "Sam Paypal" },
  },
  paypalNonBreakingSpace: {
    description: "accepts a non-breaking space before USD",
    email: {
      from: PAYPAL,
      // The space before USD is U+00A0 on purpose
      subject: "You sent a $9.99 USD payment",
      body: PAYPAL_PAYMENT_BODY,
    },
    expected: { cost: 9.99 },
  },
  paypalMoneyReceived: {
    description: "parses money received as a negative cost",
    email: {
      from: PAYPAL,
      subject: "Money is waiting for you",
      body: "Hi\n\nAccept your $20.00 USD from Bob Smith\n\nline\nline\n\nBirthday gift\nbye\n",
    },
    expected: { cost: -20, name: "Bob Smith - Birthday gift" },
  },
  paypalQuotedFrom: {
    description:
      'finds the sender in a quoted formal From ("service@paypal.com" <service@paypal.com>)',
    email: {
      from: PAYPAL_QUOTED,
      subject: "You sent a $25.50 USD payment",
      body: PAYPAL_PAYMENT_BODY,
    },
    expected: { cost: 25.5, name: "Jane Doe - Lunch money", type: "TJ Paypal" },
  },
  paypalBareQuotedFrom: {
    description:
      'finds the sender in a bare quoted From ("service@paypal.com")',
    email: {
      from: '"service@paypal.com"',
      subject: "You sent a $25.50 USD payment",
      body: PAYPAL_PAYMENT_BODY,
    },
    expected: { cost: 25.5, name: "Jane Doe - Lunch money" },
  },
  // paypal-sent-you: the bodies are made up until a real sample exists (only the subject is known)
  paypalSentYou: {
    description:
      "parses money sent to you (<Payer> sent you $X USD) as a negative cost with the note",
    email: {
      from: PAYPAL_QUOTED,
      subject: "Pat Fakename sent you $8.07 USD",
      body: "Hello\n\nPat Fakename sent you $8.07 USD\n\nNote from Pat Fakename:\nConcert tickets\n\nThanks\n",
    },
    expected: {
      cost: -8.07,
      name: "Pat Fakename - Concert tickets",
      type: "TJ Paypal",
    },
  },
  paypalSentYouNoNote: {
    description:
      "money sent to you without a note: the payer name (which may contain $ and U+00A0) from the subject",
    email: {
      from: PAYPAL,
      subject: "Dee $ollar Example sent you $1,208.07 USD",
      body: "Hello\n\nNo note here\n",
    },
    // parseFloat stops at the comma, like every other rule
    expected: { cost: -1, name: "Dee $ollar Example" },
  },
  paypalUnknownSubject: {
    description:
      "returns an empty receipt (no cost, no name) for unknown PayPal subjects",
    email: {
      from: PAYPAL,
      subject: "Update to your account",
      body: "nothing",
    },
    expected: null,
  },
  paypalForwarded: {
    description: "parses a payment forwarded from the forward address",
    email: {
      from: FWD,
      subject: "Fwd: You sent a $25.50 USD payment",
      body:
        "FYI\n\n---------- Forwarded message ---------\n" +
        "From: PayPal <service@paypal.com>\nDate: Mon, Nov 27, 2023 at 9:52 PM\n" +
        "Subject: You sent a $25.50 USD payment\nTo: Sam <fwd@example.com>\n\n" +
        PAYPAL_PAYMENT_BODY,
    },
    expected: { ...FORWARDED_PAYPAL_EXPECTED, type: "Sam Paypal" },
  },
  paypalForwardedJapanese: {
    description: "parses a forwarded email with the Japanese forwarding header",
    email: {
      from: FWD,
      subject: "Fwd: You sent a $25.50 USD payment",
      body:
        "メモ\n\n転送されたメッセージ:\n\n*差出人:* PayPal <service@paypal.com>\n" +
        "*日時:* 2024年3月3日\n*宛先:* fwd@example.com\n*件名:* *You sent*\n" +
        PAYPAL_PAYMENT_BODY,
    },
    expected: FORWARDED_PAYPAL_EXPECTED,
  },
  paypalForwardedBegin: {
    description:
      "parses a forwarded email with the 'Begin forwarded message:' header",
    email: {
      from: FWD,
      subject: "Fwd: You sent a $25.50 USD payment",
      body:
        "Hi\n\nBegin forwarded message:\n\n*From:* PayPal <service@paypal.com>\n" +
        "*Date:* December 11, 2023\n*To:* Sam <fwd@example.com>\n*Subject:* *You sent*\n\n" +
        PAYPAL_PAYMENT_BODY,
    },
    expected: FORWARDED_PAYPAL_EXPECTED,
  },

  venmoPayment: {
    description: "parses a payment with a note",
    email: {
      from: VENMO,
      subject: "You paid Jane Doe $15.00",
      body: "You paid\nJane Doe\n$15.00\nPizza night\nTransaction details\n",
    },
    expected: { cost: 15, name: "Jane Doe - Pizza night", type: "TJ Venmo" },
  },
  venmoChargeCompleted: {
    description: "parses a completed charge request",
    email: {
      from: VENMO,
      subject: "You completed Bob's $30.25 charge request",
      body: "You charged\nBob\n$30.25\nRent share\nmore\n",
    },
    expected: { cost: 30.25, name: "Bob - Rent share" },
  },
  venmoReceived: {
    description:
      "parses money received as negative cost; name is the sender from the subject",
    email: {
      from: VENMO,
      subject: "Alice paid you $40.00",
      body: "Alice paid\nyou\n$40.00\nConcert tickets\nmore\n",
    },
    expected: { cost: -40, name: "Alice - Concert tickets" },
  },
  venmoForwardedReceived: {
    description:
      "parses a forwarded Venmo received payment with the sender name",
    email: {
      from: FWD,
      subject: "Fwd: Alice paid you $40.00",
      body:
        "---------- Forwarded message ---------\nFrom: Venmo <venmo@venmo.com>\n" +
        "Date: Mon\nSubject: Alice paid you $40.00\nTo: Sam <fwd@example.com>\n\n" +
        "Alice paid\nyou\n$40.00\nConcert tickets\nmore\n",
    },
    expected: { cost: -40, name: "Alice - Concert tickets", type: "Sam Venmo" },
  },
  venmoForwarded: {
    description: "parses a forwarded Venmo payment",
    email: {
      from: FWD,
      subject: "Fwd: You paid Jane Doe $15.00",
      body:
        "---------- Forwarded message ---------\nFrom: Venmo <venmo@venmo.com>\n" +
        "Date: Mon\nSubject: You paid Jane Doe $15.00\nTo: Sam <fwd@example.com>\n\n" +
        "You paid\nJane Doe\n$15.00\nPizza night\nmore\n",
    },
    expected: { cost: 15, name: "Jane Doe - Pizza night", type: "Sam Venmo" },
  },
} satisfies Record<string, ReceiptFixture>;

/**
 * Fallback-branch fixtures the main fixtures do not cover. Pinned on the original regex parser
 * (the oracle); the rules engine must reproduce them. Names start with `edge`.
 */
export const receiptEdgeFixtures = {
  edgeChaseRefundNoMerchant: {
    description: "refund without a merchant line has no name",
    email: {
      from: CHASE,
      subject: "You have a $45.00 credit pending on your credit card",
      body: "no merchant line",
    },
    expected: { cost: -45, name: undefined, type: "Credit" },
  },
  edgeChaseGasNoMerchant: {
    description: "gas station without a merchant line is not a receipt",
    email: {
      from: CHASE,
      subject: "You used your card at a gas station",
      body: "nothing",
    },
    expected: null,
  },
  edgePaypalSentNoBody: {
    description: "payment sent with an unparsable body keeps cost but no name",
    email: {
      from: PAYPAL,
      to: "someone@example.com",
      subject: "You sent a $25.50 USD payment",
      body: "nothing",
    },
    expected: { cost: 25.5, name: undefined, type: "TJ Paypal" },
  },
  edgePaypalReceivedNoBody: {
    description: "money received with an unparsable body is not a receipt",
    email: {
      from: PAYPAL,
      subject: "Money is waiting for you",
      body: "nothing",
    },
    expected: null,
  },
  edgeVenmoPaidNoBody: {
    description: "venmo payment with an empty body has no note",
    email: {
      from: VENMO,
      subject: "You paid Jane Doe $15.00",
      body: "",
    },
    expected: { cost: 15, name: "Jane Doe" },
  },
  edgeVenmoToForwarder: {
    description: "venmo sent to the forward address uses ForwardName",
    email: {
      from: VENMO,
      to: "Sam <fwd@example.com>",
      subject: "You paid Jane Doe $15.00",
      body: "You paid\nJane Doe\n$15.00\nPizza night\nmore\n",
    },
    expected: { type: "Sam Venmo" },
  },
  edgeVenmoToForwarderQuoted: {
    description: "venmo sent to a quoted forward address uses ForwardName",
    email: {
      from: VENMO,
      to: '"fwd@example.com" <fwd@example.com>',
      subject: "You paid Jane Doe $15.00",
      body: "You paid\nJane Doe\n$15.00\nPizza night\nmore\n",
    },
    expected: { type: "Sam Venmo" },
  },
  edgeForwardedQuotedFrom: {
    description: "a forwarded header with a quoted formal From still parses",
    email: {
      from: FWD,
      subject: "Fwd: You sent a $25.50 USD payment",
      body:
        '---------- Forwarded message ---------\nFrom: "service@paypal.com" <service@paypal.com>\n' +
        "Date: Mon\nSubject: You sent a $25.50 USD payment\nTo: Sam <fwd@example.com>\n\n" +
        PAYPAL_PAYMENT_BODY,
    },
    expected: { cost: 25.5, type: "Sam Paypal" },
  },
  edgeUnknownSender: {
    description:
      "a message from an unknown sender is not a receipt (e.g. a reply in a group thread)",
    email: UNKNOWN_SENDER_REPLY,
    expected: null,
  },
  edgeForwarderReply: {
    description:
      "a message from the forward address without a forward prefix is not a receipt",
    email: FORWARDER_REPLY,
    expected: null,
  },
  edgeVenmoToUnparsable: {
    description: "venmo with an unparsable 'to' falls back to TJ",
    email: {
      from: VENMO,
      to: "undisclosed-recipients",
      subject: "You paid Jane Doe $15.00",
      body: "You paid\nJane Doe\n$15.00\nPizza night\nmore\n",
    },
    expected: { type: "TJ Venmo" },
  },
  edgeForwardedChase: {
    description: "a forwarded Chase email is not forwardable",
    email: {
      from: FWD,
      subject: "Fwd: You made a $1.00 transaction with X",
      body:
        "---------- Forwarded message ---------\nFrom: Chase <no.reply.alerts@chase.com>\n" +
        "Date: Mon\nSubject: You made a $1.00 transaction with X\nTo: Sam <fwd@example.com>\n\nbody\n",
    },
    expected: null,
  },
  edgeForwardedTentativeFrom: {
    description: "a forwarded header with a bare address still parses",
    email: {
      from: FWD,
      subject: "Fwd: You sent a $25.50 USD payment",
      body:
        "---------- Forwarded message ---------\nFrom: service@paypal.com\n" +
        "Date: Mon\nSubject: You sent a $25.50 USD payment\nTo: Sam <fwd@example.com>\n\n" +
        PAYPAL_PAYMENT_BODY,
    },
    expected: { cost: 25.5, type: "Sam Paypal" },
  },
  edgeCostNotNumber: {
    description: "a non-numeric cost parses as NaN",
    email: {
      from: CHASE,
      subject: "You made a $abc transaction with Shop",
      body: "irrelevant",
    },
    expected: { name: "Shop", costIsNaN: true },
  },
  edgeZeroCostNoName: {
    description: "a zero refund with no merchant has no information",
    email: {
      from: CHASE,
      subject: "You have a $0.00 credit pending on your credit card",
      body: "no merchant line",
    },
    expected: null,
  },
} satisfies Record<string, ReceiptFixture>;

export type ReceiptFixtureName = keyof typeof receiptFixtures;

/** Multi-message threads (messages in order) for the labeling tests */
export const receiptThreadFixtures = {
  /** Only unknown-sender and forwarder replies: ignored (Receipts/Ignored, stays in the inbox) */
  unknownSendersOnly: [
    { ...UNKNOWN_SENDER_REPLY, date: FIXTURE_DATE },
    { ...FORWARDER_REPLY, date: FIXTURE_DATE },
  ],
  /** A receipt plus an unknown-sender reply: valid (receipt recorded, archived) */
  receiptWithUnknownReply: [
    { ...receiptFixtures.paypalSentYou.email, date: FIXTURE_DATE },
    { ...UNKNOWN_SENDER_REPLY, date: FIXTURE_DATE },
  ],
  /**
   * The forwarder forwards a non-receipt from someone with no rules, who then replies. The query
   * finds it (forward address + "Fwd" subject); it is ignored
   */
  forwardedFromUnknownSender: [
    {
      from: FWD,
      subject: "Fwd: Dinner Friday?",
      body:
        "See below\n\n---------- Forwarded message ---------\nFrom: Pat Friend <pat.friend@example.net>\n" +
        "Date: Thu\nSubject: Dinner Friday?\nTo: Sam <fwd@example.com>\n\nAnyone free?\n",
      date: FIXTURE_DATE,
    },
    { ...UNKNOWN_SENDER_REPLY, date: FIXTURE_DATE },
  ],
} satisfies Record<string, EmailFixture[]>;

/** A Chase purchase email for tests that just need a receipt thread */
export function chasePurchaseEmail(
  cost: string,
  merchant: string,
  date: string,
): EmailFixture {
  return {
    from: CHASE,
    subject: `You made a $${cost} transaction with ${merchant}`,
    body: "",
    date: new Date(date),
  };
}

/**
 * The fake inbox of the demo backend: fixtures whose parse results are one receipt each (the
 * Gas station fixture is used with a body that has no trailing spaces), plus one non-receipt
 */
export const demoInbox: Record<string, EmailFixture> = {
  chasePurchase: {
    ...receiptFixtures.chasePurchase.email,
    subject: "You made a $18.42 transaction with Demo Cafe",
  },
  chaseGasStation: {
    ...receiptFixtures.chaseGasStation.email,
    body: "Hello\nMerchant   Fake Fuel 42\n\nMore text\n",
  },
  venmoPayment: receiptFixtures.venmoPayment.email,
  paypalPaymentSent: receiptFixtures.paypalPaymentSent.email,
  chaseStatement: receiptFixtures.chaseStatement.email,
};
