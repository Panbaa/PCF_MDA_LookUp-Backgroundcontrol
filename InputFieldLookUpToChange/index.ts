import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class InputFieldLookUpToChange implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    // Reference to the control container HTMLDivElement
	// This element contains all elements of our custom control example
	private _container: HTMLDivElement;

	// Reference to ComponentFramework Context object
	private _context: ComponentFramework.Context<IInputs>;

	// PCF framework delegate which will be assigned to this object which would be called whenever any update happens
	private _notifyOutputChanged: () => void;

	// Containers to store and display raw lookup data for both properties
	private _inputData: HTMLLabelElement;

	// Values to be filled based on control context, passed as arguments to lookupObjects API
	private _entityType = "";
	private _defaultViewId = "";

	// Used to store necessary data for a single lookup entity selected during runtime
	private _selectedItem: ComponentFramework.LookupValue;
    private _colorInRgb: string;
    root = document.documentElement;

	constructor() { }

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */


	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
		this._container = container;
		this._notifyOutputChanged = notifyOutputChanged;

		// Cache information necessary for lookupObjects API based on control context
		this._entityType = context.parameters.controlValue.getTargetEntityType();
		this._defaultViewId = context.parameters.controlValue.getViewId();

		const contentContainer = document.createElement("div");
        contentContainer.setAttribute("id", "lookUp");

		const lookUpItem = document.createElement("div");
		lookUpItem.setAttribute("id", "lookUpItem");

		// Add element to display raw entity data
		this._inputData = document.createElement("label");
        this._inputData.setAttribute("id", "lookUpLabel");
		this._inputData.textContent = "";
		lookUpItem.appendChild(this._inputData);

        // Add element to remove entity data
        const deleteButton = document.createElement("button");
        deleteButton.setAttribute("id", "lookUpDeleteButton");
        deleteButton.onclick = this.clearLookupField.bind(this, (value, update = true) => {
            this._selectedItem = value;
        });
        lookUpItem.appendChild(deleteButton);
		
		contentContainer.appendChild(lookUpItem);

        // Add element to display search-icon
        const lookUpIcon = document.createElement("div");
        lookUpIcon.setAttribute("id", "lookUpIcon");
        contentContainer.appendChild(lookUpIcon);

		// Add button to trigger lookupObjects API for the primary lookup property
		const lookupObjectsButton = document.createElement("button");
        lookupObjectsButton.setAttribute("id", "lookUpButton");
		lookupObjectsButton.onclick = this.performLookupObjects.bind(this, this._entityType, this._defaultViewId, (value, update = true) => {
			this._selectedItem = value;
		});
		contentContainer.appendChild(lookupObjectsButton);

		this._container.appendChild(contentContainer);

        this.root.style.setProperty('--backgroundColor', this._colorInRgb);
	}

	/**
	 * Helper to utilize the lookupObjects API. Gets whatever lookup record is selected and allows
	 * the control to use the received data.
	 * @param entityType The entity logical name bound to the target lookup property
	 * @param viewId The viewId bound to the target lookup property
	 * @param setSelected Specified function to set the selected lookup value
	 */
	private performLookupObjects(entityType: string, viewId: string, setSelected: (value: ComponentFramework.LookupValue, update?: boolean) => void): void {
		// Used cached values from lookup parameter to set options for lookupObjects API
		const lookupOptions = {
			defaultEntityType: entityType,
			defaultViewId: viewId,
			allowMultiSelect: false,
			entityTypes: [entityType],
			viewIds: [viewId]
		};

		this._context.utils.lookupObjects(lookupOptions).then((success) => {
			if (success && success.length > 0) {
				// Cache the necessary information for the newly selected entity lookup
				const selectedReference = success[0];
				const selectedLookupValue: ComponentFramework.LookupValue = {
					id: selectedReference.id,
					name: selectedReference.name,
					entityType: selectedReference.entityType
				};

				// Update the primary or secondary lookup property
				setSelected(selectedLookupValue);

				// Trigger a control update
				this._notifyOutputChanged();
			} else {
				setSelected({} as ComponentFramework.LookupValue);
				this._notifyOutputChanged();
			}
		}, (error) => {
			console.log(error);
		});
	}

    private clearLookupField(setSelected: (value: ComponentFramework.LookupValue, update?: boolean) => void) {
		console.log("Der Delete-Button wurde geklickt.");
        // Set the selected value to an empty object to clear the lookup field
        setSelected({} as ComponentFramework.LookupValue);

		this._inputData.textContent = "";
        
        // Trigger a control update
        this._notifyOutputChanged();
    }

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Update the main text field of the control to contain the raw data of the entity selected via lookup

		const lookupValue: ComponentFramework.LookupValue = context.parameters.controlValue.raw[0];
		this._context = context;
		let propertyValue = lookupValue ? lookupValue.name : "";

		// Ensure propertyValue is not undefined
		if (propertyValue === undefined) {
			propertyValue = "";
		}

		this._inputData.textContent = propertyValue;

		this._colorInRgb = context.parameters.ColorInRGB.raw || "";
		console.log("Color: " + this._colorInRgb + "inputColor: " + context.parameters.ColorInRGB.raw);
        this.root.style.setProperty('--backgroundColor', this._colorInRgb);
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
	 */
	public getOutputs(): IOutputs {
		// Send the updated selected lookup item back to the ComponentFramework, based on the currently selected item
		return { controlValue: [this._selectedItem] };
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// no-op: method not leveraged by this example custom control
	}
}
