/**
 * A class for generating and manipulating Google Documents.
 */
class DocumentGenerator
{
	/**
	 * @private
	 * @type {GoogleAppsScript.Document.Document}
	 */
	templateDocument;

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
	 * @private
	 * @type {string}
	 */
	escapedPlaceholderStart;

	/**
	 * @private
	 * @type {string}
	 */
	escapedPlaceholderEnd;

	/**
	 * @param {string} docUrlOrId The URL or ID of the Google Document.
	 * @param {string=} placeholderStart Optional. The starting delimiter for placeholders. Defaults to '<<'.
	 * @param {string=} placeholderEnd Optional. The ending delimiter for placeholders. Defaults to '>>'.
	 */
	constructor(docUrlOrId, placeholderStart = '<<', placeholderEnd = '>>')
	{
		if (docUrlOrId.startsWith('https://'))
		{
			this.templateDocument = DocumentApp.openByUrl(docUrlOrId);
		}
		else
		{
			this.templateDocument = DocumentApp.openById(docUrlOrId);
		}

		this.placeholderStart = placeholderStart;
		this.placeholderEnd = placeholderEnd;

		this.escapedPlaceholderStart = this._escapeRegExp(placeholderStart);
		this.escapedPlaceholderEnd = this._escapeRegExp(placeholderEnd);

		this.placeholderRegex = new RegExp(this.escapedPlaceholderStart + '(.*?)' + this.escapedPlaceholderEnd, 'g');
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
	 * Returns the currently managed template document.
	 * @return {GoogleAppsScript.Document.Document} The Google Document.
	 */
	getTemplateDocument()
	{
		return this.templateDocument;
	}

	/**
	 * Returns the ID of the currently managed template document.
	 * @return {string} The ID of the Google Document.
	 */
	getTemplateDocumentId()
	{
		return this.templateDocument.getId();
	}

	/**
	 * Replaces all occurrences of a specific placeholder key with a given value in the managed document.
	 * @param {string} key The placeholder key (e.g., "name" for "<<name>>").
	 * @param {string} value The value to replace the placeholder with.
	 */
	replacePlaceholder(key, value)
	{
		const body = this.templateDocument.getBody();
		const specificPlaceholderRegex = new RegExp(
			this.escapedPlaceholderStart +
			this._escapeRegExp(key) +
			this.escapedPlaceholderEnd,
			'g'
		);
		body.replaceText(specificPlaceholderRegex, value);
	}

	/**
	 * Collects all unique placeholder keys from the document.
	 * @return {Array<string>} An array of unique placeholder keys.
	 */
	getPlaceholders()
	{
		const body = this.templateDocument.getBody();
		const text = body.getText();
		const matches = text.matchAll(this.placeholderRegex);
		return [...new Set(Array.from(matches, match => match[1]))];
	}

	/**
	 * Saves and closes the currently managed document.
	 */
	saveAndCloseDocument()
	{
		this.templateDocument.saveAndClose();
	}
}