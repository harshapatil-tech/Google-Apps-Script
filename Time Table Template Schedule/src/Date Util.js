class DateUtil {

  // /** @private */
  // static DEFAULT_LOCALE = 'en-US';

  static getBaseDate(overrideDate) {
    return overrideDate instanceof Date ? overrideDate : new Date(2025, 11, 28);
  }

  /**
   * Gets the current year (four digits).
   * @param {Date=} date
   * @returns {number}
   */
  static getCurrentYear (date) {
    return DateUtil.getBaseDate(date).getFullYear();
  }

  /**
   * Gets the current month number (1-12).
   * @param {Date=} date
   * @returns {number}
   */
  static getCurrentMonthNum (date) {
    return DateUtil.getBaseDate(date).getMonth() + 1;
  }

  /**
   * Gets the current month name ("long" or "short")
   * @param {Date=} date
   * @param { "long" | "short" } format
   * @returns {string}
   */
  static getCurrentMonthStr (date, format="short") {
    return DateUtil.getBaseDate(date).toLocaleDateString("en-GB", {month:format})
  }

  /**
   * Returns the Date representing the first day of the next month.
   * @param {Date=} date
   * @returns {Date}
   */
  static getNextMonthDate(date) {
    const d = DateUtil.getBaseDate(date);
    return new Date(d.getFullYear(), d.getMonth() + 1, 1);
  }

  /**
   * Gets the month number (1-12) for the next month.
   * @param {Date=} date
   * @returns {number}
   */
  static getNextMonthNumber(date) {
    return DateUtil.getNextMonthDate(date).getMonth() + 1;
  }

  /**
   * Gets the year for the next month (e.g., year increment if December rolls over).
   * @param {Date=} date
   * @returns {number}
   */
  static getNextYear(date) {
    return DateUtil.getNextMonthDate(date).getFullYear();
  }

   /**
   * Gets the next month's name in specified format.
   * @param {'short'|'long'|'numeric'} format
   * @param {Date=} date
   * @param {string=} locale
   * @returns {string}
   */
  static getNextMonthName(format, date) {
    return DateUtil.getNextMonthDate(date).toLocaleDateString('en-GB',{ month: format });
  }



  /**
   * Core formatter: takes a JS Date and a pattern, returns the formatted string.
   * @private
   */
  static _formatKeyFromDate(d, pattern, locale = 'en-GB') {
    const fullYear   = d.getFullYear();
    const shortYear  = String(fullYear).slice(-2);
    const monthNum   = String(d.getMonth() + 1).padStart(2, '0');
    const monthShort = d.toLocaleString(locale, { month: 'short' });
    const monthLong  = d.toLocaleString(locale, { month: 'long' });

    return pattern
      .replace(/{YYYY}/g, fullYear)
      .replace(/{MONTH}/g, monthLong)
      .replace(/{MON}/g, monthShort)
      .replace(/{MM}/g, monthNum)
      .replace(/{M}/g, String(d.getMonth() + 1))
      .replace(/{YY}/g, shortYear);
  }

  /**
   * Formats any date according to pattern.
   */
  static formatMonthKey(date, pattern = '{MM}{MON}-{YY}', locale) {
    const base = DateUtil.getBaseDate(date);
    return DateUtil._formatKeyFromDate(base, pattern, locale);
  }

  /**
   * Shorthand for the “current” month.
   */
  static getMonthKey(date, pattern = "{MM}{MON}-{YY}", locale) {
    return DateUtil.formatMonthKey(date, pattern, locale);
  }

  /**
   * ***NEW***: Shorthand for the *next* month.
   */
  static getNextMonthKey(date, pattern = "{MM}{MON}-{YY}", locale) {
    // pull your “base date” then bump it one month:
    const next = DateUtil.getNextMonthDate(date);
    return DateUtil._formatKeyFromDate(next, pattern, locale);
  }

  /**
   * Finds the Monday on or before the first day of the given month/year.
   * @param {number} year - Four-digit year.
   * @param {number} month - Month number (1-12).
   * @returns {Date}
   * @throws {RangeError} If inputs are invalid.
   */
  static firstWeekMonday(year, month) {
    if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) {
      throw new RangeError('Invalid year or month');
    }
    const first = new Date(year, month - 1, 1);
    const offset = (first.getDay() + 6) % 7;
    first.setDate(first.getDate() - offset);
    return first;
  }


}


function test() {
  console.log(DateUtil.getMonthKey(
  new Date(2025, 11, 11),"{MM} {MONTH}'{YY}"));
}


