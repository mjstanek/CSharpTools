using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace DataverseMappingRemover.Services
{
    public class MappingInspectorService
    {
        private readonly IOrganizationService _organizationService;
        public MappingInspectorService(IOrganizationService organizationService)
        {
            _organizationService = organizationService ?? throw new ArgumentNullException(nameof(organizationService));
        }

        public async Task<List<MappingResult>> FindAttributeMappingsAsync(string sourceEntityLogicalName,
            string sourceAttributeLogicalName, CancellationToken cancellationToken = default)
        {
            if (sourceEntityLogicalName.Contains(" "))
            {
                throw new ArgumentException("Entity logical name must not contain spaces.",
                    nameof(sourceEntityLogicalName));
            }
            if (sourceAttributeLogicalName.Contains(" "))
            {
                throw new ArgumentException("Attribute logical name must not contain spaces.",
                    nameof(sourceAttributeLogicalName));
            }
            if (string.IsNullOrWhiteSpace(sourceEntityLogicalName))
            {
                throw new ArgumentException("Entity logical name is required.",
                    nameof(sourceEntityLogicalName));
            }
            else if (string.IsNullOrWhiteSpace(sourceAttributeLogicalName))
            {
                throw new ArgumentException("Attribute logical name is required.",
                    nameof(sourceAttributeLogicalName));
            }

            sourceEntityLogicalName = sourceEntityLogicalName.Trim().ToLowerInvariant();
            sourceAttributeLogicalName = sourceAttributeLogicalName.Trim().ToLowerInvariant();

            var results = new List<MappingResult>();

            await Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();

                var sourceQuery = new Microsoft.Xrm.Sdk.Query.QueryExpression("attributemap")
                {
                    ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("attributemapid","sourceattributename",
                    "targetattributename", "entitymapid")
                };

                sourceQuery.Criteria.AddCondition("sourceattributename", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal,
                    sourceAttributeLogicalName);

                var linkEntityMap = new LinkEntity(
                    "attributemap",
                    "entitymap",
                    "entitymapid",
                    "entitymapid",
                    JoinOperator.Inner)
                {
                    Columns = new ColumnSet("sourceentityname", "targetentityname"),
                    EntityAlias = "em"
                };

                linkEntityMap.LinkCriteria.AddCondition("sourceentityname", ConditionOperator.Equal,
                    sourceEntityLogicalName);

                sourceQuery.LinkEntities.Add(linkEntityMap);

                var sourceResults = _organizationService.RetrieveMultiple(sourceQuery);

                cancellationToken.ThrowIfCancellationRequested();

                foreach (var entity in sourceResults.Entities)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var mappingResult = new MappingResult
                    {
                        SourceEntity = entity.GetAttributeValue<AliasedValue>("em.sourceentityname")?.Value.ToString(),
                        SourceAttribute = entity.GetAttributeValue<string>("sourceattributename"),
                        TargetEntity = entity.GetAttributeValue<AliasedValue>("em.targetentityname")?.Value.ToString(),
                        TargetAttribute = entity.GetAttributeValue<string>("targetattributename"),
                        AttributeMapId = entity.Id,
                        EntityMapId = entity.GetAttributeValue<EntityReference>("entitymapid").Id,
                        Direction = "Source to Target"
                    };
                    results.Add(mappingResult);
                }

                cancellationToken.ThrowIfCancellationRequested();

                var targetQuery = new Microsoft.Xrm.Sdk.Query.QueryExpression("attributemap")
                {
                    ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("attributemapid", "sourceattributename",
                    "targetattributename", "entitymapid")
                };
                
                targetQuery.Criteria.AddCondition("targetattributename", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal,
                    sourceAttributeLogicalName);

                var targetLinkEntityMap = new LinkEntity(
                    "attributemap",
                    "entitymap",
                    "entitymapid",
                    "entitymapid",
                    JoinOperator.Inner)
                {
                    Columns = new ColumnSet("sourceentityname", "targetentityname"),
                    EntityAlias = "em"
                };

                targetLinkEntityMap.LinkCriteria.AddCondition("targetentityname", ConditionOperator.Equal,
                    sourceEntityLogicalName);

                targetQuery.LinkEntities.Add(targetLinkEntityMap);

                var targetResults = _organizationService.RetrieveMultiple(targetQuery);

                cancellationToken.ThrowIfCancellationRequested();

                foreach (var entity in targetResults.Entities)
                {

                    cancellationToken.ThrowIfCancellationRequested();

                    var mappingResult = new MappingResult
                    {
                        SourceEntity = entity.GetAttributeValue<AliasedValue>("em.sourceentityname")?.Value.ToString(),
                        SourceAttribute = entity.GetAttributeValue<string>("sourceattributename"),
                        TargetEntity = entity.GetAttributeValue<AliasedValue>("em.targetentityname")?.Value.ToString(),
                        TargetAttribute = entity.GetAttributeValue<string>("targetattributename"),
                        AttributeMapId = entity.Id,
                        EntityMapId = entity.GetAttributeValue<EntityReference>("entitymapid").Id,
                        Direction = "Target to Source"
                    };
                    results.Add(mappingResult);
                }

            }, cancellationToken);

            var deduped = new List<MappingResult>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var result in results)
            {
                var key = $"{result.SourceEntity}|{result.SourceAttribute}|{result.TargetEntity}|" +
                    $"{result.TargetAttribute}|{result.EntityMapId}|{result.Direction}";
                if (!seen.Contains(key))
                {
                    seen.Add(key);
                    deduped.Add(result);
                }
            }

            return deduped;

        }
    }
}
