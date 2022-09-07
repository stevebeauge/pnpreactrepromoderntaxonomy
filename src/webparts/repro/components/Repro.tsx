import * as React from 'react';


import { IModernTaxonomyPickerProps, ModernTaxonomyPicker } from '@pnp/spfx-controls-react';

import { BaseComponentContext } from '@microsoft/sp-component-base';
import { TextField } from 'office-ui-fabric-react';
import { useState } from 'react';

type ReproProps = {
  context: BaseComponentContext
};

type PartialTermInfo = IModernTaxonomyPickerProps["initialValues"];

const Repro = ({ context }: ReproProps): React.ReactElement<ReproProps> => {

  const [termSetId, setTermSetId] = useState<string | undefined>(undefined);

  const [selectedTerms, setSelectedTerms] = useState<PartialTermInfo | undefined>(undefined);

  return (
    <>
      <div>
        <TextField value={termSetId} onChange={(evt, newValue) => setTermSetId(newValue)} />
      </div>
      <div>
        {termSetId && (
          <ModernTaxonomyPicker
            termSetId={termSetId}
            panelTitle="My term set"
            label="Modern taxonmy picker"
            context={context}
            initialValues={selectedTerms}
            onChange={setSelectedTerms}
          />
        )}
      </div> <div>
        {selectedTerms && (
          <ul>
            {selectedTerms.map(t => (
              <li key={t.id}>{t.labels[0].name}</li>
            ))
            }
          </ul>
        )
        }
      </div>
    </>
  );
}

export { Repro, ReproProps };
