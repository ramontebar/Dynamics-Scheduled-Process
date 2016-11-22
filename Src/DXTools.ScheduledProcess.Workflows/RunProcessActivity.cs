using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.ScheduledProcess.Workflows
{
    public class RunProcessActivity : BaseCodeActivity
    {
        #region In Arguments

        [Input("Process Name")]
        [Default("")]
        public InArgument<String> ProcessName { get; set; }

        [Input("Process Type")]
        [AttributeTarget("dxtools_scheduledprocess", "dxtools_processtype")]
        public InArgument<OptionSetValue> ProcessType { get; set; }

        [Input("Execute On")]
        [AttributeTarget("dxtools_scheduledprocess","dxtools_executeon")]
        public InArgument<OptionSetValue> ExecuteOn { get; set; }

        [Input("Record ID")]
        public InArgument<String> RecordID { get; set; }

        #endregion

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            Guid processid = RetrieveProcess(executionContext);

            ProcessTypeEnum processType = GetProcessType(executionContext);

            switch (processType)
            {
                case ProcessTypeEnum.Workflow:
                    RunWorkflow(processid, executionContext);
                    break;
                case ProcessTypeEnum.Action:
                    throw new NotImplementedException();    
            }
        }

        private void RunWorkflow(Guid processid, CodeActivityContext executionContext)
        {
            ExecuteOnEnum executeOn = GetExecuteOn(executionContext);

            switch (executeOn)
            {
                case ExecuteOnEnum.Global:
                    ExecuteWorkflow(processid, this.WorkflowContext.PrimaryEntityId);
                    break;
                case ExecuteOnEnum.SingleRecord:
                    Guid recordID = GetRecordID(executionContext);
                    ExecuteWorkflow(processid, recordID);
                    break;
                case ExecuteOnEnum.Query:
                    throw new NotImplementedException();
            }           
        }

        private void ExecuteWorkflow(Guid processID, Guid recordID)
        {
            ExecuteWorkflowRequest executeWorkflowRequest = new ExecuteWorkflowRequest();
            executeWorkflowRequest.WorkflowId = processID;
            executeWorkflowRequest.EntityId = recordID;
            ExecuteWorkflowResponse response = this.OrganizationService.Execute(executeWorkflowRequest) as ExecuteWorkflowResponse;
            if (response != null)
                TraceService.Trace("Global Workflow has been executed correctly with ID: '{0}'.", response.Id);
            else
                TraceService.Trace("Global Workflow has been executed correctly but response is NULL.");
        }

        private Guid GetRecordID(CodeActivityContext executionContext)
        {
            throw new NotImplementedException();
        }

        private ExecuteOnEnum GetExecuteOn(CodeActivityContext executionContext)
        {
            if (this.ExecuteOn == null)
                throw new InvalidPluginExecutionException("ExecuteOn argument cannot be null");

            OptionSetValue executeOnOptionSetValue = this.ProcessType.Get<OptionSetValue>(executionContext);
            if (executeOnOptionSetValue == null)
                throw new InvalidPluginExecutionException("ExecuteOn cannot be null");

            ExecuteOnEnum executeOn;

            string executeOnStringValue = executeOnOptionSetValue.Value.ToString();
            TraceService.Trace("Process Execute On value has been resolved correctly: '{0}'.", executeOnStringValue);

            if (Enum.TryParse<ExecuteOnEnum>(executeOnStringValue, out executeOn))
                return executeOn;
            else
                throw new InvalidPluginExecutionException(string.Format("Unexpected Execute On value '{0}'", executeOnStringValue));

        }

        private ProcessTypeEnum GetProcessType(CodeActivityContext executionContext)
        {
            if (this.ProcessType == null)
                throw new InvalidPluginExecutionException("Process Type argument cannot be null");

            OptionSetValue processTypeOptionSetValue = this.ProcessType.Get<OptionSetValue>(executionContext);
            if(processTypeOptionSetValue==null)
                throw new InvalidPluginExecutionException("Process Type cannot be null");

            ProcessTypeEnum processType;

            string procesTypeStringValue = processTypeOptionSetValue.Value.ToString();
            TraceService.Trace("Process Type value has been resolved correctly: '{0}'.", procesTypeStringValue);

            if (Enum.TryParse<ProcessTypeEnum>(procesTypeStringValue, out processType))
                return processType;
            else
                throw new InvalidPluginExecutionException(string.Format("Unexpected Process Type value '{0}'", procesTypeStringValue));

        }

        private Guid RetrieveProcess(CodeActivityContext executionContext)
        {
            if (this.ProcessName == null)
                throw new InvalidPluginExecutionException("Process Name argument cannot be null");

            String processName = this.ProcessName.Get<String>(executionContext);
            if(string.IsNullOrEmpty(processName))
                throw new InvalidPluginExecutionException("Process Name cannot be null or empty");

            QueryExpression query = new QueryExpression("workflow");
            query.ColumnSet = new ColumnSet(new string[] { "workflowid" });
            query.Criteria.AddCondition("name", ConditionOperator.Equal, processName);
            query.Criteria.AddCondition("type", ConditionOperator.Equal, 1); //Workflow type equal to Definition

            EntityCollection processCollection = this.OrganizationService.RetrieveMultiple(query);

            if (processCollection.Entities == null)
                throw new InvalidPluginExecutionException("No Entities were found collection response");

            if (processCollection.Entities.Count == 0)
                throw new InvalidPluginExecutionException(string.Format("No Processes were found with name '{0}'", processName));

            if (processCollection.Entities.Count > 1)
                throw new InvalidPluginExecutionException(string.Format("More than one process was found with name '{0}'", processName));

            Guid processID = processCollection.Entities[0].Id;
            TraceService.Trace("The process '{0}' was retrieved correctly with ID '{1}'", processName, processID);

            return processID;
        }
    }
}
