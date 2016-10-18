using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.Note
{
	public class MoveNote : CodeActivity
	{
		[RequiredArgument]
		[Input("Note To Move")]
		[ReferenceTarget("annotation")]
		public InArgument<EntityReference> NoteToMove { get; set; }

		[Input("Record Dynamic Url")]
		[RequiredArgument]
		public InArgument<string> RecordUrl { get; set; }

		[Output("Was Note Moved")]
		public OutArgument<bool> WasNoteMoved { get; set; }

		protected override void Execute(CodeActivityContext executionContext)
		{
			ITracingService tracer = executionContext.GetExtension<ITracingService>();
			IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
			IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

			try
			{
				EntityReference noteToMove = NoteToMove.Get(executionContext);
				string recordUrl = RecordUrl.Get<string>(executionContext);

				var dup = new DynamicUrlParser(recordUrl);

				string newEntityLogical = dup.GetEntityLogicalName(service);

				Entity note = GetNote(service, noteToMove.Id);
				if (note.GetAttributeValue<EntityReference>("objectid").Id == dup.Id && note.GetAttributeValue<EntityReference>("objectid").LogicalName == newEntityLogical)
				{
					WasNoteMoved.Set(executionContext, false);
					return;
				}

				Entity updateNote = new Entity("annotation");
				updateNote.Id = noteToMove.Id;
				updateNote["objectid"] = new EntityReference(newEntityLogical, dup.Id);

				service.Update(updateNote);

				WasNoteMoved.Set(executionContext, true);
			}
			catch (Exception e)
			{
				throw new InvalidPluginExecutionException(e.Message);
			}
		}

		private Entity GetNote(IOrganizationService service, Guid noteId)
		{
			return service.Retrieve("annotation", noteId, new ColumnSet("objectid", "objecttypecode"));
		}
	}
}
