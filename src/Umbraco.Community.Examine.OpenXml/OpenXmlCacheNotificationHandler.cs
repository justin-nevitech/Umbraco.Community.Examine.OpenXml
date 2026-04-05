using Microsoft.Extensions.Logging;
using Umbraco.Cms.Core;
using Umbraco.Cms.Core.Cache;
using Umbraco.Cms.Core.Events;
using Umbraco.Cms.Core.Notifications;
using Umbraco.Cms.Core.Services;
using Umbraco.Cms.Core.Services.Changes;
using Umbraco.Cms.Core.Sync;
using Umbraco.Extensions;

namespace Umbraco.Community.Examine.OpenXml
{
    public class OpenXmlCacheNotificationHandler : INotificationHandler<MediaCacheRefresherNotification>
    {
        private readonly IMediaService _mediaService;
        private readonly OpenXmlIndexPopulator _openXmlIndexPopulator;
        private readonly IRuntimeState _runtimeState;
        private readonly ILogger<OpenXmlCacheNotificationHandler> _logger;

        public OpenXmlCacheNotificationHandler(IMediaService mediaService, OpenXmlIndexPopulator openXmlIndexPopulator, IRuntimeState runtimeState, ILogger<OpenXmlCacheNotificationHandler> logger)
        {
            _mediaService = mediaService;
            _openXmlIndexPopulator = openXmlIndexPopulator;
            _runtimeState = runtimeState;
            _logger = logger;
        }

        /// <summary>
        /// Handles the cache refresher event and updates the index.
        /// </summary>
        /// <param name="notification"></param>
        public void Handle(MediaCacheRefresherNotification notification)
        {
            // Only handle events when the site is running.
            if (_runtimeState.Level != RuntimeLevel.Run)
            {
                return;
            }

            try
            {
                if (notification.MessageType != MessageType.RefreshByPayload)
                {
                    _logger.LogWarning("Unsupported message type {MessageType} received, skipping", notification.MessageType);
                    return;
                }

                var payloads = notification.MessageObject as MediaCacheRefresher.JsonPayload[];
                if (payloads == null)
                {
                    _logger.LogWarning("Unexpected message object type {Type}, skipping", notification.MessageObject?.GetType().Name);
                    return;
                }

                foreach (var payload in payloads)
                {
                    if (payload.ChangeTypes.HasType(TreeChangeTypes.Remove))
                    {
                        _openXmlIndexPopulator.RemoveFromIndex(payload.Id);
                    }
                    else if (payload.ChangeTypes.HasType(TreeChangeTypes.RefreshAll))
                    {
                        // RefreshAll is not supported by Examine index operations.
                        // A full index rebuild should be triggered from the Examine Management dashboard instead.
                        _logger.LogDebug("RefreshAll received, skipping. Use the Examine Management dashboard to rebuild the index");
                    }
                    else // RefreshNode or RefreshBranch (maybe trashed)
                    {
                        var media = _mediaService.GetById(payload.Id);

                        if (media is null)
                        {
                            _openXmlIndexPopulator.RemoveFromIndex(payload.Id);
                            continue;
                        }

                        if (media.Trashed)
                        {
                            _openXmlIndexPopulator.RemoveFromIndex(payload.Id);
                        }
                        else
                        {
                            _openXmlIndexPopulator.AddToIndex(media);
                        }

                        // Branch
                        if (payload.ChangeTypes.HasType(TreeChangeTypes.RefreshBranch))
                        {
                            const int pageSize = 500;
                            var page = 0;
                            var total = long.MaxValue;
                            while (page * pageSize < total)
                            {
                                var descendants = _mediaService.GetPagedDescendants(media.Id, page++, pageSize, out total);
                                foreach (var descendant in descendants)
                                {
                                    if (descendant.Trashed)
                                        _openXmlIndexPopulator.RemoveFromIndex(descendant);
                                    else
                                        _openXmlIndexPopulator.AddToIndex(descendant);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error handling media cache notification");
            }
        }
    }
}
