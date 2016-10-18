using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Text;

namespace LAT.WorkflowUtilities.Note
{
	public class DeleteAttachmentByName : CodeActivity
	{
		[RequiredArgument]
		[Input("Note With Attachment To Remove")]
		[ReferenceTarget("annotation")]
		public InArgument<EntityReference> NoteWithAttachment { get; set; }

		[RequiredArgument]
		[Input("File Name With Extension")]
		public InArgument<string> FileName { get; set; }

		[RequiredArgument]
		[Input("Add Delete Notice As Text?")]
		[Default("false")]
		public InArgument<bool> AppendNotice { get; set; }

		[Output("Number Of Attachments Deleted")]
		public OutArgument<int> NumberOfAttachmentsDeleted { get; set; }

		protected override void Execute(CodeActivityContext executionContext)
		{
			ITracingService tracer = executionContext.GetExtension<ITracingService>();
			IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
			IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
			IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

			try
			{
				EntityReference noteWithAttachment = NoteWithAttachment.Get(executionContext);
				string fileName = FileName.Get(executionContext);
				bool appendNotice = AppendNotice.Get(executionContext);

				Entity note = GetNote(service, noteWithAttachment.Id);
				if (!CheckForAttachment(note))
					return;

				StringBuilder notice = new StringBuilder();
				int numberOfAttachmentsDeleted = 0;

				if (String.Equals(note.GetAttributeValue<string>("filename"), fileName, StringComparison.CurrentCultureIgnoreCase))
				{
					numberOfAttachmentsDeleted++;

					if (appendNotice)
						notice.AppendLine("Deleted Attachment: " + note.GetAttributeValue<string>("filename") + " " +
						                  DateTime.Now.ToShortDateString());

					UpdateNote(service, note, notice.ToString());
				}

				NumberOfAttachmentsDeleted.Set(executionContext, numberOfAttachmentsDeleted);
			}
			catch (Exception ex)
			{
				tracer.Trace("Exception: {0}", ex.ToString());
			}
		}

		private static bool CheckForAttachment(Entity note)
		{
			object oIsAttachment;
			bool hasValue = note.Attributes.TryGetValue("isdocument", out oIsAttachment);
			if (!hasValue)
				return false;

			return (bool)oIsAttachment;
		}

		private static Entity GetNote(IOrganizationService service, Guid noteId)
		{
			return service.Retrieve("annotation", noteId, new ColumnSet("filename", "isdocument", "notetext"));
		}

		private void UpdateNote(IOrganizationService service, Entity note, string notice)
		{
			Entity updateNote = new Entity("annotation");
			updateNote.Id = note.Id;
			if (!string.IsNullOrEmpty(notice))
			{
				string newText = note.GetAttributeValue<string>("notetext");
				if (!string.IsNullOrEmpty(newText))
					newText += "\r\n";

				updateNote["notetext"] = newText + notice;
			}
			updateNote["isdocument"] = false;
			updateNote["filename"] = null;
			updateNote["documentbody"] = null;
			updateNote["filesize"] = null;

			service.Update(updateNote);
		}
	}
}
