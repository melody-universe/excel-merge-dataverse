using ExcelMerge.Library;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMerge.CodeActivities
{
    public class MergeExcelFiles : CodeActivity
    {
        static byte[] Base64Template = Convert.FromBase64String(
            "<REDACTED>"
        );

        [RequiredArgument]
        [Input("Id of the record to fatten all attachments on")]
        public new InArgument<string> Id { get; set; }

        [Output("Flattened Excel file")]
        public OutArgument<string> OutputFile { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            var service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);

            var query = new QueryExpression("annotation")
            {
                ColumnSet = new ColumnSet("documentbody"),
                Criteria = new FilterExpression(
                    LogicalOperator.And)
            };
            query.Criteria.AddCondition("filesize", ConditionOperator.GreaterThan, 0);
            query.Criteria.AddCondition(
                "objectid",
                ConditionOperator.Equal,
                Guid.Parse(Id.Get(context)));
            query.Criteria.AddCondition(
                "filename",
                ConditionOperator.EndsWith,
                ".xlsx");
            query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
            var results = service.RetrieveMultiple(query);
            var documents = results.Entities.Select(
                annotation => Convert.FromBase64String(
                    annotation.GetAttributeValue<string>("documentbody")
                )
            );
            
            var outputContents = Excel.Merge(
                Base64Template,
                documents,
                new int[]
                {
                    5, 6, 7, 8
                }
            );
            
            OutputFile.Set(context, Convert.ToBase64String(outputContents));
        }
    }
}
