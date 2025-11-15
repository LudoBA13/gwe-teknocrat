/**
 * A class for generating and manipulating Google Documents.
 */
class DocumentGenerator
{
	/**
	 * @private
	 * @type {GoogleAppsScript.Document.Document}
	 */
	document;

	/**
	 * @private
	 * @type {string}
	 */
	placeholderStart;

	/**
	 * @private
	 * @type {string}
	 */
	placeholderEnd;

	/**
	 * @private
	 * @type {RegExp}
	 */
	placeholderRegex;

	/**
	 * @param {string} docUrlOrId The URL or ID of the Google Document.
	 * @param {string=} placeholderStart Optional. The starting delimiter for placeholders. Defaults to '<<'.
	 * @param {string=} placeholderEnd Optional. The ending delimiter for placeholders. Defaults to '>>'.
	 */
	constructor(docUrlOrId, placeholderStart = '<<', placeholderEnd = '>>')
	{
		if (docUrlOrId.startsWith('https://'))
		{
			this.document = DocumentApp.openByUrl(docUrlOrId);
		}
		else
		{
			this.document = DocumentApp.openById(docUrlOrId);
		}

		this.placeholderStart = placeholderStart;
		this.placeholderEnd = placeholderEnd;

		const escapedPlaceholderStart = this._escapeRegExp(placeholderStart);
		const escapedPlaceholderEnd = this._escapeRegExp(placeholderEnd);

		this.placeholderRegex = new RegExp(escapedPlaceholderStart + '(.*?)' + escapedPlaceholderEnd, 'g');
	}

	/**
	 * Escapes special characters in a string for use in a regular expression.
	 * @param {string} str The string to escape.
	 * @return {string} The escaped string.
	 * @private
	 */
	_escapeRegExp(str)
	{
		return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
	}

	/**
	 * Returns the currently managed document.
	 * @return {GoogleAppsScript.Document.Document} The Google Document.
	 */
	getDocument()
	{
		return this.document;
	}

	/**
	 * Returns the ID of the currently managed document.
	 * @return {string} The ID of the Google Document.
	 */
	getDocumentId()
	{
		return this.document.getId();
	}

	/**
	 * Replaces all occurrences of a specific placeholder key with a given value in the managed document.
	 * @param {string} key The placeholder key (e.g., "name" for "<<name>>").
	 * @param {string} value The value to replace the placeholder with.
	 */
	replacePlaceholder(key, value)
	{
		const body = this.document.getBody();
		const specificPlaceholderRegex = new RegExp(
			this._escapeRegExp(this.placeholderStart) +
			this._escapeRegExp(key) +
			this._escapeRegExp(this.placeholderEnd),
			'g'
		);
		body.replaceText(specificPlaceholderRegex, value);
	}

	/**
	 * Saves and closes the currently managed document.
	 */
	saveAndCloseDocument()
	{
		this.document.saveAndClose();
	}
}