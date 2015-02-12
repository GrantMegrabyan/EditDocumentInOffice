var DocumentEditing = DocumentEditing || {};

DocumentEditing.OfficeDocumentEditor = function()
{
    var winFirefoxPluginContainerId = '_winFirefoxPluginContainer';
    var pluginsHolder = 'body';

    var getPlugin = function()
    {
        if ('ActiveXObject' in window)
        {
            return getSharepointPlugin();
        }
        else
        {
            return getWinFirefoxPlugin();
        }
    };

    var getSharepointPlugin = function()
    {
        if (typeof ($(pluginsHolder).data('sharepointPlugin')) === 'undefined')
        {
            var plugin = new ActiveXObject('SharePoint.OpenDocuments.3');
            $(pluginsHolder).data('sharepointPlugin', plugin);
        }

        return $(pluginsHolder).data('sharepointPlugin');
    };

    var getWinFirefoxPlugin = function()
    {
        var pluginContainer = getOrCreateContainer(winFirefoxPluginContainerId);

        var domObject = $('object', pluginContainer)[0];
        if (typeof (domObject) == 'undefined')
        {
            var $domObject = $('<object />');
            $domObject.attr('type', 'application/x-sharepoint');
            $domObject.css('visibility', 'hidden');
            $domObject.attr('width', '0');
            $domObject.attr('height', '0');

            pluginContainer.append($domObject);

            domObject = $domObject[0];
        }

        return domObject;
    };

    var getOrCreateContainer = function(containerId)
    {
        var container = $('#' + containerId);
        if (container.length != 0)
        {
            return container;
        }

        container = $('<div />');
        container.attr('id', containerId);
        container.css('width', '0');
        container.css('height', '0');
        container.css('position', 'absolute');
        container.css('overflow', 'hidden');
        container.css('top', '1000px');
        container.css('left', '1000px');

        $(pluginsHolder).append(container);
        return container;
    };

    var editDocument = function(url)
    {
        var plugin = getPlugin();

        if (!('EditDocument' in plugin))
        {
            throw 'EditDocument is not supported';
        }

        if (!plugin.EditDocument(url))
        {
            alert('Cannot edit file');
        }
    };

    var isEditDocumentSupported = function()
    {
        var plugin = getPlugin();

        return 'EditDocument' in plugin;
    };

    return {
        EditDocument: editDocument,
        IsSupported: isEditDocumentSupported
    }
}();