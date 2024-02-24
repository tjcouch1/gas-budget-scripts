namespace Budgeting {
  // #region regexp helper functions and stuff

  /**
   * Provides a comparison of each character to its unicode code point
   * @param str string to analyze
   * @returns pretty-printed stringified array of arrays of character to unicode code point
   */
  function getUnicodeAnalysis(str: string) {
    return JSON.stringify(
      str.split("").map((char) => [char, char.codePointAt(0)?.toString(16)]),
      undefined,
      2
    );
  }

  /**
   * RegExp that matches to `message.getFrom()` to get the actual email address
   *
   * Ex: Chase <no.reply.alerts@chase.com>
   *
   * Named groups: `email`
   */
  const emailFromRegExp = /<(?<email>.+)>/;

  /**
   * Get a RegExp modified to match to forwarded email subjects
   * @param regex regex to modify
   * @returns RegExp matching to forwarded emails with the same subject
   */
  function getForwardSubjectRegExp(regex: RegExp) {
    // Take off /^ from front and / from back
    return new RegExp(`^Fwd: ${regex.toString().slice(2, -1)}`);
  }

  /**
   * Get the original email address from a forwarded email
   * @param message message for which to get the forwarded email
   * @returns email address that sent the original email before being forwarded
   */
  function getForwardEmailAddress(
    message: GoogleAppsScript.Gmail.GmailMessage
  ) {
    const matches =
      /[Ff]orwarded message.*\s*\r?\n\r?\n?\*?From:\*?[ \u00a0](?<email>.+)\r?\n/.exec(
        message.getPlainBody()
      );
    if (!matches || matches.length <= 0 || !matches.groups)
      throw new Error(
        `Could not get 'from' email address from forwarded email with subject '${message.getSubject()}'`
      );

    // May be the email address or may be the gmail formatted "from" with angle brackets
    // so return the actual email address
    const tentativeFrom = matches.groups.email;
    const fromMatches = emailFromRegExp.exec(tentativeFrom);
    return fromMatches && fromMatches.length > 0 && fromMatches.groups
      ? fromMatches.groups.email
      : tentativeFrom;
  }

  /**
   * Get the original email plain body from a forwarded email
   * @param message message for which to get the forwarded info
   * @returns original plain body for the forwarded email
   */
  function getForwardPlainBody(message: GoogleAppsScript.Gmail.GmailMessage) {
    const plainBody = message.getPlainBody();

    /* Ex:
---------- Forwarded message ---------
From: American Eagle <ae@notifications.ae.com>
Date: Mon, Nov 27, 2023 at 9:52 PM
Subject: Order Confirmed! #0153262784
To: Keilah <keilahfok@gmail.com>



Thanks, ...
    */
    let fwInd = plainBody.indexOf("---------- Forwarded message ---------");
    if (fwInd >= 0)
      // Skip the first 5 lines
      return plainBody.substring(fwInd).split("\n").slice(5).join("\n");

    /* Ex:

Return receipt


Begin forwarded message:

*From:* Bare Necessities <DoNotReply@barenecessities.com>
*Date:* December 11, 2023 at 9:09:41 PM CST
*To:* Keilah Couch <keilahfok@gmail.com>
*Subject:* *Bare Necessities Order # BN23689811 RETURNED Item Confirmation*

Email contents ...
    */
    fwInd = plainBody.indexOf("Begin forwarded message:");
    if (fwInd >= 0)
      // Skip the first 6 lines
      return plainBody.substring(fwInd).split("\n").slice(6).join("\n");
    return plainBody;
  }

  /**
   * Test if a part of a message matches the provided `RegExp` and get the pieces of message info contained if so
   * @param messagePart subject or body of a message
   * @param regExp `RegExp` to `exec` against the `messagePart`. This `RegExp` must contained named groups `cost` and `name`
   */
  function getMessagePartInfo(
    messagePart: string,
    regExp: RegExp
  ):
    | {
        cost: number | undefined;
        name: string | undefined;
        details: string | undefined;
      }
    | undefined {
    const matches = regExp.exec(messagePart);
    if (matches && matches.length > 0) {
      return {
        cost: matches.groups?.cost
          ? parseFloat(matches.groups.cost)
          : undefined,
        name: matches.groups?.name,
        details: matches.groups?.details,
      };
    }
    return undefined;
  }

  // #endregion

  // #region chase regexp

  /**
   * RegExp pattern matching to Chase's receipt email subjects
   *
   * Named groups: `cost`, `name`
   */
  const chaseSubjectReceiptRegExp =
    /^Your \$(?<cost>.+) transaction with (?<name>.+)$/;
  /**
   * RegExp pattern matching to Chase's refund receipt email subjects
   *
   * Named groups: `cost`
   */
  const chaseSubjectRefundRegExp =
    /^You have a \$(?<cost>.+) credit pending on your credit card$/;
  /**
   * RegExp pattern matching to Chase's gas receipt email subjects
   */
  const chaseSubjectGasRegExp = /^You used your card at a gas station$/;
  /**
   * RegExp pattern matching to Chase's refund and gas receipt email plain body
   *
   * Named groups: `name`
   */
  const chaseBodyMerchantRegExp = /\r?\nMerchant\s+(?<name>.+)\s+\r?\n/;

  // #endregion

  // #region paypal regexp

  /** Receipt from which we will check for venmo receipts */
  const paypalReceiptEmailAddress = "service@paypal.com";

  /**
   * RegExp pattern matching to paypal's receipt email subjects
   *
   * Named groups: `cost`
   */
  const paypalSubjectReceiptRegExp =
    /^You sent a \$(?<cost>.+)[\s\u00a0]USD payment$/;
  /**
   * RegExp pattern matching to paypal's receipt email subjects forwarded from others
   *
   * Named groups: `cost`
   */
  const paypalForwardSubjectReceiptRegExp = getForwardSubjectRegExp(
    paypalSubjectReceiptRegExp
  );
  /**
   * RegExp pattern matching to Paypal's receipt email plain body to get the recipient name
   *
   * Named groups: `name`, `details`
   */
  const paypalBodyReceiptRecipientRegExp =
    /You sent \$.+ to (?<name>.+)\r?\n\r?\n.*\r?\n.*\r?\n\r?\n(?<details>.+)\r?\n/;
  /**
   * RegExp pattern matching to paypal's received email subjects
   *
   * Named groups: `cost`
   */
  const paypalSubjectReceivedRegExp = /^Money is waiting for you$/;
  /**
   * RegExp pattern matching to paypal's received email subjects forwarded from others
   *
   * Named groups: `cost`
   */
  const paypalForwardSubjectReceivedRegExp = getForwardSubjectRegExp(
    paypalSubjectReceivedRegExp
  );
  /**
   * RegExp pattern matching to Paypal's received email plain body to get the sender name,
   * transaction details, and transaction cost
   *
   * Named groups: `cost`, `name`, `details`
   */
  const paypalBodyReceivedNoteRegExp =
    /Accept your \$(?<cost>.+)[\s\u00a0]USD from (?<name>.+)\r?\n\r?\n.*\r?\n.*\r?\n\r?\n(?<details>.+)\r?\n/;

  // #endregion

  // #region venmo regexp

  /** Receipt from which we will check for venmo receipts */
  const venmoReceiptEmailAddress = "venmo@venmo.com";

  /**
   * RegExp pattern matching to Venmo's receipt email subjects
   *
   * Named groups: `cost`, `name`
   */
  const venmoSubjectReceiptRegExp = /^You paid (?<name>.+) \$(?<cost>.+)$/;
  /**
   * RegExp pattern matching to Venmo's receipt email subjects forwarded from others
   *
   * Named groups: `cost`, `name`
   */
  const venmoForwardSubjectReceiptRegExp = getForwardSubjectRegExp(
    venmoSubjectReceiptRegExp
  );
  /**
   * RegExp pattern matching to Venmo's receipt email plain body to get the transaction note
   *
   * Named groups: `details`
   */
  const venmoBodyReceiptNoteRegExp =
    /paid\s*.*\s*\s*.*\s*\r?\n(?<details>.+)\r?\n/;
  /**
   * RegExp pattern matching to Venmo's completed charge email subjects
   *
   * Named groups: `cost`, `name`
   */
  const venmoSubjectChargeRegExp =
    /^You completed (?<name>.+)'s \$(?<cost>.+) charge request$/;
  /**
   * RegExp pattern matching to Venmo's completed charge email subjects forwarded from others
   *
   * Named groups: `cost`, `name`
   */
  const venmoForwardSubjectChargeRegExp = getForwardSubjectRegExp(
    venmoSubjectChargeRegExp
  );
  /**
   * RegExp pattern matching to Venmo's charge email plain body to get the transaction details
   *
   * Named groups: `name`
   */
  const venmoBodyChargeNoteRegExp =
    /charged\s*.*\s*\s*.*\s*\r?\n(?<details>.+)\r?\n/;
  /**
   * RegExp pattern matching to Venmo's received receipt email subjects
   *
   * Named groups: `cost`, `name`
   */
  const venmoSubjectReceivedRegExp = /^(?<name>.+) paid you \$(?<cost>.+)$/;
  /**
   * RegExp pattern matching to Venmo's received receipt email subjects forwarded from others
   *
   * Named groups: `cost`, `name`
   */
  const venmoForwardSubjectReceivedRegExp = getForwardSubjectRegExp(
    venmoSubjectReceivedRegExp
  );
  /**
   * RegExp pattern matching to Venmo's received email plain body to get the transaction details
   *
   * Named groups: `name`
   */
  const venmoBodyReceivedNoteRegExp =
    /paid\s*.*\s*\s*.*\s*\r?\n(?<details>.+)\r?\n/;

  // #endregion

  /**
   * Get information about a paypal receipt email
   * @param message email message for which to get receipt info
   * @param isForwarded whether the email was forwarded from others
   * @param typePrefix prefix to add to the receipt type (and a space added)
   * @returns Receipt information
   */
  function getReceiptInfoPaypal(
    message: GoogleAppsScript.Gmail.GmailMessage,
    isForwarded: boolean,
    typePrefix: string
  ): ReceiptInfo {
    const subject = message.getSubject();
    // Try to get the cost and name for the message
    let cost: number | undefined;
    let name: string | undefined;
    const type = `${typePrefix} Paypal`;

    // Test if it is a normal paypal receipt
    let matches = getMessagePartInfo(
      subject,
      isForwarded
        ? paypalForwardSubjectReceiptRegExp
        : paypalSubjectReceiptRegExp
    );
    if (matches) {
      cost = matches.cost;
      matches = getMessagePartInfo(
        isForwarded ? getForwardPlainBody(message) : message.getPlainBody(),
        paypalBodyReceiptRecipientRegExp
      );
      if (matches) name = `${matches.name} - ${matches.details}`;
    } else {
      // Test if it is a paypal received receipt
      matches = getMessagePartInfo(
        subject,
        isForwarded
          ? paypalForwardSubjectReceivedRegExp
          : paypalSubjectReceivedRegExp
      );
      if (matches) {
        matches = getMessagePartInfo(
          isForwarded ? getForwardPlainBody(message) : message.getPlainBody(),
          paypalBodyReceivedNoteRegExp
        );
        if (matches) {
          cost = matches.cost! * -1;
          name = `${matches.name} - ${matches.details}`;
        }
      }
    }

    return new Budgeting.ReceiptInfo(message, cost, name, undefined, type);
  }

  /**
   * Get information about a venmo receipt email
   * @param message email message for which to get receipt info
   * @param isForwarded whether the email was forwarded from others
   * @param typePrefix prefix to add to the receipt type (and a space added)
   * @returns Receipt information
   */
  function getReceiptInfoVenmo(
    message: GoogleAppsScript.Gmail.GmailMessage,
    isForwarded: boolean,
    typePrefix: string
  ): ReceiptInfo {
    const subject = message.getSubject();
    // Try to get the cost and name for the message
    let cost: number | undefined;
    let name: string | undefined;
    const type = `${typePrefix} Venmo`;

    // Test if it is a normal venmo receipt
    let matches = getMessagePartInfo(
      subject,
      isForwarded ? venmoForwardSubjectReceiptRegExp : venmoSubjectReceiptRegExp
    );
    if (matches) {
      cost = matches.cost;
      name = matches.name;
      matches = getMessagePartInfo(
        isForwarded ? getForwardPlainBody(message) : message.getPlainBody(),
        venmoBodyReceiptNoteRegExp
      );
      if (matches) name = `${name} - ${matches.details}`;
    } else {
      // Test if it is a venmo charge receipt
      matches = getMessagePartInfo(
        subject,
        isForwarded ? venmoForwardSubjectChargeRegExp : venmoSubjectChargeRegExp
      );
      if (matches) {
        cost = matches.cost;
        name = matches.name;
        matches = getMessagePartInfo(
          isForwarded ? getForwardPlainBody(message) : message.getPlainBody(),
          venmoBodyChargeNoteRegExp
        );
        if (matches) name = `${name} - ${matches.details}`;
      } else {
        // Test if it is a venmo received receipt
        matches = getMessagePartInfo(
          subject,
          isForwarded
            ? venmoForwardSubjectReceivedRegExp
            : venmoSubjectReceivedRegExp
        );
        if (matches) {
          cost = matches.cost! * -1;
          matches = getMessagePartInfo(
            isForwarded ? getForwardPlainBody(message) : message.getPlainBody(),
            venmoBodyReceivedNoteRegExp
          );
          if (matches) name = `${name} - ${matches.details}`;
        }
      }
    }

    return new Budgeting.ReceiptInfo(message, cost, name, undefined, type);
  }

  /**
   * Map of email "from" address to functions to get `ReceiptInfo` about the receipt from a Gmail `message`
   * @param message Gmail `message` to translate into a `ReceiptInfo`
   * @returns `ReceiptInfo` for the Gmail `message`. Not a receipt message if cost and name are not filled in
   */
  const getReceiptInfoMap: {
    [addressFrom: string]:
      | ((message: GoogleAppsScript.Gmail.GmailMessage) => ReceiptInfo)
      | undefined;
  } = {
    "no.reply.alerts@chase.com": function getReceiptInfoChase(
      message: GoogleAppsScript.Gmail.GmailMessage
    ): ReceiptInfo {
      const subject = message.getSubject();
      // Try to get the cost and name for the message
      let cost: number | undefined;
      let name: string | undefined;
      const type = "Credit";

      // Test if it is a normal chase receipt
      let matches = getMessagePartInfo(subject, chaseSubjectReceiptRegExp);
      if (matches) {
        cost = matches.cost;
        name = matches.name;
      } else {
        // Test if it is a chase return receipt
        matches = getMessagePartInfo(subject, chaseSubjectRefundRegExp);
        if (matches) {
          cost = matches.cost! * -1;
          matches = getMessagePartInfo(
            message.getPlainBody(),
            chaseBodyMerchantRegExp
          );
          if (matches) name = matches.name;
        } else {
          // Test if it is a chase gas receipt
          matches = getMessagePartInfo(subject, chaseSubjectGasRegExp);
          if (matches) {
            matches = getMessagePartInfo(
              message.getPlainBody(),
              chaseBodyMerchantRegExp
            );
            if (matches) name = matches.name;
          }
        }
      }

      return new Budgeting.ReceiptInfo(message, cost, name, undefined, type);
    },
    [paypalReceiptEmailAddress]: (message) =>
      getReceiptInfoPaypal(message, false, "TJ"),
    [venmoReceiptEmailAddress]: (message) =>
      getReceiptInfoVenmo(message, false, "TJ"),
    "keilahfok@gmail.com": (message) => {
      const forwardedFrom = getForwardEmailAddress(message);

      if (forwardedFrom === paypalReceiptEmailAddress)
        return getReceiptInfoPaypal(message, true, "Keilah");
      if (forwardedFrom === venmoReceiptEmailAddress)
        return getReceiptInfoVenmo(message, true, "Keilah");

      // Couldn't process the email. Just return empty message so we make a note
      return new Budgeting.ReceiptInfo(
        message,
        undefined,
        undefined,
        undefined,
        undefined
      );
    },
  };

  /**
   * Compares receiptInfos by date in ascending order
   *
   * Used in [`Array.prototype.sort`](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)
   */
  function compareReceiptInfosByDateAscending(
    a: ReceiptInfoBase,
    b: ReceiptInfoBase
  ) {
    if (a.date > b.date) return 1;
    if (Util.areDatesEqual(a.date, b.date)) return 0;
    return -1;
  }

  /** Information for threads retrieved and receipts found in those threads */
  export class ThreadList {
    /** Array of receipt-related thread info */
    public readonly threadInfos: ThreadInfo[] = [];

    #allReceiptInfosCache: ReceiptInfoBase[] | undefined;
    /**
     * Returns an array of all receipt email info - flat mapped from threadInfos and contains thread notes and errors baked into the receipts
     *
     * TODO: Move note and error calculation stuff into ThreadInfo?
     */
    public getAllReceiptInfos(): ReceiptInfoBase[] {
      if (this.#allReceiptInfosCache) return this.#allReceiptInfosCache;

      this.#allReceiptInfosCache = this.threadInfos.flatMap((threadInfo) => {
        const receiptInfos: ReceiptInfoBase[] = [...threadInfo.receiptInfos];

        // Add an empty receipt for an error line if the thread has no receipts
        if (receiptInfos.length === 0) {
          receiptInfos.push(
            new Budgeting.EmptyReceipt(
              threadInfo.thread.getLastMessageDate(),
              threadInfo.thread
            )
          );
        }

        // Add errors to the first receipt
        if (threadInfo.errors.length > 0)
          receiptInfos[0].errorMessage = threadInfo.errors.join(
            "\n\n~~~~~~~~~~~~~~~~~ NEXT ERROR ~~~~~~~~~~~~~~~~~\n\n"
          );

        // Add notes to the first receipt
        if (threadInfo.notes.length > 0)
          receiptInfos[0].note = threadInfo.notes.join(
            "\n\n~~~~~~~~~~~~~~~~~ NEXT NOTE ~~~~~~~~~~~~~~~~~\n\n"
          );

        return receiptInfos;
      });
      this.#allReceiptInfosCache.sort(compareReceiptInfosByDateAscending);

      return this.#allReceiptInfosCache;
    }
    /**
     *
     * @param threads array of GmailThreads from which to derive receipts
     */
    constructor(threads: GoogleAppsScript.Gmail.GmailThread[]) {
      // Map email info into receipt infos
      /** All receipt email info */
      threads.forEach((thread) => {
        const threadInfo = new Budgeting.ThreadInfo(thread);

        try {
          // Try getting receipt info from each message in the thread
          const messages = thread.getMessages();

          messages.forEach((message) => {
            try {
              // Get 'from' email address out of the message
              // Ex: Chase <no.reply.alerts@chase.com>
              const messageFrom = message.getFrom();
              const matches = emailFromRegExp.exec(messageFrom);
              if (!matches || matches.length <= 0 || !matches.groups)
                throw new Error(
                  `Could not get 'from' email address from ${messageFrom}`
                );
              // Actual 'from' email address
              const from = matches.groups.email;

              const getReceiptInfo = getReceiptInfoMap[from];
              if (!getReceiptInfo)
                throw new Error(
                  `No function to handle receipt message from ${from}`
                );

              const receipt = getReceiptInfo(message);

              if (!receipt.cost && !receipt.name) {
                // This message is not a receipt. Add a note about it
                threadInfo.notes.push(
                  `Message is not a receipt:\nFrom: ${from}\nSubject: ${message.getSubject()}\nDate: ${message.getDate()}\nThread ID: ${thread.getId()}\n280 Chars of Plain Body:\n${message
                    .getPlainBody()
                    ?.substring(0, 280)}`
                );
              } else {
                // We have a receipt (or a blank receipt with a note). Return receiptInfo
                threadInfo.receiptInfos.push(receipt);
              }
            } catch (e) {
              threadInfo.errors.push(
                `Error while processing message ${message.getId()} with subject ${message.getSubject()} from date ${message.getDate()} on thread ${thread.getId()}. Skipping marking as processed. ${e}`
              );
            }
          });
        } catch (e) {
          threadInfo.errors.push(
            `Error while processing thread with ID ${thread.getId()}. Skipping marking as processed. ${e}`
          );
        }

        // Save the thread info if there were any receipts in the thread or there was an error
        if (threadInfo.hasInformation()) this.threadInfos.push(threadInfo);
        else if (threadInfo.notes.length > 0)
          // Log any notes we are throwing out so we know
          Logger.log(
            `Ignoring thread ${thread.getId()} that has notes but no relevant info. Notes: ${
              threadInfo.notes
            }`
          );
      });
    }
  }
}
