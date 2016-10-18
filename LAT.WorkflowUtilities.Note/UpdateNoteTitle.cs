using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.Note
{
	public class UpdateNoteTitle : CodeActivity
	{
		[RequiredArgument]
		[Input("Note To Update")]
		[ReferenceTarget("annotation")]
		public InArgument<EntityReference> NoteToUpdate { get; set; }

		[RequiredArgument]
		[Input("New Title")]
		public InArgument<string> NewTitle { get; set; }

		[Output("Updated Title")]
		public OutArgument<string> UpdatedTitle { get; set; }

		protected override void Execute(CodeActivityContext executionContext)
		{
			ITracingService tracer = executionContext.GetExtension<ITracingService>();
			IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
			IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

			try
			{
				EntityReference noteToUpdate = NoteToUpdate.Get(executionContext);
				string newTitle = NewTitle.Get(executionContext);

				Entity note = new Entity("annotation");
				note.Id = noteToUpdate.Id;
				note["subject"] = newTitle;
				service.Update(note);

				UpdatedTitle.Set(executionContext, newTitle);
			}
			catch (Exception e)
			{
				throw new InvalidPluginExecutionException(e.Message);
			}
		}
	}
}
