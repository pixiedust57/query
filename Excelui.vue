/* eslint-disable */
<template>
  <div>
    <q-btn-group
      spread
      v-if="getProposaldata && getProposaldata.proposalRefId && interCoComId !== company.id"
      style="max-width:88.8vw"
    >
      <q-btn
        size="xs"
        icon="timeline"
        :label="jExcelOptions.sheetName"
        :color="selectedtab === 'primary' ? 'accent' : 'grey-5'"
        @click="selectedtab = 'primary'"
      />
      <q-btn
        size="xs"
        icon="visibility"
        :label="refoptions.sheetName"
        :color="selectedtab === 'secondary' ? 'accent' : 'grey-5'"
        @click="selectedtab = 'secondary'"
      />
    </q-btn-group>

    <q-dialog v-model="taxes">
      <q-card style="width:600px">
        <q-card-section class="row">
          <div class="text-h6">Taxable Data </div>
          <q-space />
          <q-btn
            dense
            flat
            round
            size="xs"
            icon="close"
            v-close-popup
            class="float-right"
            @click="taxes = false"
          >
            <q-tooltip
              max-height="30%"
              max-width="25%"
              content-class="bg-white text-black"
            >
              <div class="tooltip-content">
                Close
              </div>
            </q-tooltip>
          </q-btn>
        </q-card-section>
        <q-separator class="q-mb-md" />
        <q-card-section
          v-if="offerTax.taxcharges && offerTax.taxcharges.length"
        >
          <div v-if="offerTax.taxformatlabels" class="row">
            <div
              class="q-gutter-lg row"
              v-for="format in offerTax.taxformatlabels"
              :key="format"
            >
              <q-radio
                dense
                class=" q-mr-sm"
                v-model="offerTax['taxformat']"
                :val="format"
                :label="format"
                @input="changetaxtype(offerTax['taxformat'])"
              />
            </div>
          </div>
          <q-list>
            <div
              v-for="(taxRow, index) in offerTax.taxcharges"
              :key="index"
              class="row q-gutter-sm"
            >
              <q-input
                flat
                dense
                autogrow
                borderless
                :readonly="true"
                v-model="taxRow.name"
                class="col-md-5 col-lg-5 col-sm-5 col-xs-12"
              />
              <q-input
                dense
                min="0"
                outlined
                max="100"
                type="number"
                label="Percent %"
                v-model="taxRow.percent"
                class="col-md-3 col-lg-3 col-sm-3 col-xs-12"
                v-if="Object.keys(taxRow).includes('percent')"
                :rules="[
                  val => (val >= 0 && val < 100) || 'Enter b/w 0 & 100'
                ]"
                @input="calculatedvalue(taxRow, { template: true, index })"
              />
              <q-input
                v-else-if="Object.keys(taxRow).includes('comment')"
                dense
                :label="taxRow.isamount ? 'ex Factory Work' : 'Comment'"
                v-model="taxRow.comment"
                outlined
                class="col-md-3 col-lg-3 col-sm-3 col-xs-12"
              />
              <q-space />
              <q-input
                dense
                min="0"
                max="100"
                outlined
                type="number"
                label="Price"
                :prefix="currency"
                v-model="taxRow.price"
                class="col-md-3 col-lg-3 col-sm-3 q-mb-lg col-xs-12"
                @input="calculatedvalue(taxRow, index)"
                :disable="!taxRow.comment"
              />
            </div>
          </q-list>
          <div></div>
        </q-card-section>

        <q-card-actions align="right">
          <q-btn
            no-caps
            unelevated
            size="xs"
            icon="save"
            color="primary"
            class="q-ma-xs"
            label="Save & continue"
            @click="savetax()"
          />
        </q-card-actions>
      </q-card>
    </q-dialog>

    <div class="wrapper-jexcel">
      <div
        v-show="getProposaldata && !getProposaldata.proposalRefId || selectedtab == 'primary'"
        id="spreadsheet"
        ref="spreadsheet"
        class="xl-with-custom-header offer-sheet"
      />
      <div v-if="getProposaldata && getProposaldata.proposalRefId">
        <div
          v-show="getProposaldata.proposalRefId && selectedtab === 'secondary'"
          id="spreadsheet2"
          ref="spreadsheet2"
          class="xl-with-custom-header"
        />
      </div>
      <q-dialog v-model="nsMdl" persistent>
        <q-card style="min-width: 700px">
          <q-card-section>
            <div class="text-h6 bg-grey-3">Non-Standard Detail</div>
          </q-card-section>

          <q-card-section class="q-pt-none row">
            <q-select
              dense
              ref="nstext"
              v-model="nsText"
              outlined
              input-debounce="0"
              :options="nonStandardList"
              label="Type and hit enter to add NS Item"
              class="col-5"
              clearable
              use-input
              option-value="name"
              option-label="name"
              @filter="filternsList"
              @keyup.enter="ns = false"
              @input="setnsDependency"
              @new-value="
                (val, done) =>
                  createNSvalue(val, done, nsColumnNameTarget.dependenciesId)
              "
            />
            <q-space />

            <q-select
              v-if="
                nsColumnNameTarget.dependenciesId &&
                  nsColumnNameTarget.dependenciesId.length
              "
              dense
              v-model="nsDependentvalue"
              outlined
              input-debounce="0"
              :options="nsDependendlist"
              class="col-5"
              clearable
              use-input
              multiple
              use-chips
              option-value="id"
              option-label="name"
              label="Dependent value"
              @keyup.enter="ns = false"
              @filter="filternsdependent"
            />
          </q-card-section>

          <q-card-actions align="right" class="text-primary">
            <q-btn
              unelevated
              flat
              size="xs"
              label="Cancel"
              class="btnCancle"
              @click="
                nsOptions = []
                nsDependentvalue = null
                nsText = null
              "
              v-close-popup
            />
            <q-btn
              size="xs"
              unelevated
              color="primary"
              label="Add Ns details"
              @click="saveNs"
            />
            <!-- class="btnSave" -->
          </q-card-actions>
        </q-card>
      </q-dialog>
      <q-dialog
        v-model="secondDialog"
        persistent
        transition-show="scale"
        transition-hide="scale"
      >
        <q-card class="bg-teal text-white" style="width: 300px">
          <q-card-section>
            <div class="text-h6">Persistent</div>
          </q-card-section>

          <q-card-section>Click/Tap on the backdrop.</q-card-section>

          <q-card-actions align="right" class="bg-white text-teal">
            <q-btn size="xs" flat label="OK" v-close-popup unelevated />
          </q-card-actions>
        </q-card>
      </q-dialog>
      <q-dialog v-model="comment"  seamless>
        <q-card dense class="q-py-none" style="max-width: 1500px; width: 800px">
        <div class="row">
        <q-toolbar dense class="bg-accent text-white text-h6 col-11"> {{ currentcell.title }} Data</q-toolbar>
        <q-btn
              dense
              flat
              square
              size="xs"
              color="grey-11"
              icon="close"
              class="q-pa-xs float-right bg-accent col-1"
              @click="saveDellData('close')"
            >
              <q-tooltip
                max-height="100%"
                max-width="40%"
                content-class="bg-white text-black"
                >Close</q-tooltip
              >
            </q-btn>
            </div>
          <q-card-section class="q-pt-none q-py-none">
            <q-form class="q-gutter-y-xs">
              <q-btn
               v-if="(company.code !== '71' && currentcell.datasheet)"
                dense
                size="xs"
                color="primary"
                label="Technical datasheet"
                @click="technicalData = true"
              />
            <div v-if="company.code === '71'">
              <div class="row q-mt-sm" >
                <span class="col-3"> Short Description </span>
                <q-input
                  dense
                  disable
                  type="text"
                  class="col-8 text-caption"
                  v-model="formdata['desc']"
                />
              </div>
              <div class="row" >
                <span class="col-3"> Long Description </span>
                <q-input
                  dense
                  disable
                  autogrow
                  type="textarea"
                  class="col-9 text-caption"
                  v-model="formdata['descLg']"
                />
              </div>
            </div>
          <div>
            <div v-if="company.code !== '71'">
              <div class="row q-mt-sm" >
                <span class="col-4">Item Code</span>
                <q-select
                  dense
                  outlined
                  clearable
                  use-input
                  class="col-8"
                  option-value="id"
                  input-debounce="0"
                  v-model="itemCode"
                  option-label="code"
                  :options="listOfItemcode"
                  new-value-mode="add-unique"
                  @input="onItemCodeInput"
                  @filter="filterItemCodeList"
                  @new-value="createItemCode"
                >
                  <template v-slot:no-option>
                    <q-item>
                      <q-item-section class="text-grey">
                        No results
                      </q-item-section>
                    </q-item>
                  </template>
                </q-select>
              </div>

              <div class="row q-mt-sm">
                <span class="col-4">Description</span>
                <!-- EN 10204 3.1 Material & Functional Test Reports: -->
                <q-input
                  dense
                  outlined
                  type="textarea"
                  class="col-8"
                  v-model="formdata['desc']"
                />
              </div>
              <div class="row q-mt-sm">
                <span class="col-4"> Price</span>
                <!-- Recommended Spares for 1yr Operation: -->
                <q-input
                  dense
                  min="0"
                  outlined
                  type="number"
                  class="col-8"
                  :prefix="currency"
                  v-model="formdata['price']"
                  :rules="[val => val >= 0 || 'Value should be positive']"
                  @input="handlePriceInput"
                />
              </div>
             <div class="row">
                <span class="col-4"> Qty</span>
                 <q-input
                    dense
                    v-model.number="formdata['qty']"
                    type="number"
                    label="Qty"
                    min="1"
                    :rules="[val => val >= 1 || 'Value should be positive']"
                    class="col-1"
                    outlined
                  />
              </div>
              <div  class="q-gutter-y-md">
                <div class="row" v-if="formdata['multiplyval']">
                  <span class="col-4"> Multiplier</span>
                  <!-- Recommended Spares for 1yr Operation: -->
                  <q-input
                    dense
                    v-model.number="formdata['multiplier']"
                    type="number"
                    min="0"
                    :rules="[val => val >= 0 || 'Value should be positive']"
                    class="col-8"
                    outlined
                  />
                </div>
                <div class="row">
                  <span class="col-4">Final {{ currentcell.title }} Price</span>
                  <!-- Recommended Spares for 1yr Operation: -->
                  <q-input
                    dense
                    disable
                    outlined
                    type="number"
                    class="col-8"
                    :prefix="currency"
                    v-model="getfinalprice"
                  />
                </div>
              </div>
            </div>
            </div>
              <div class="row" v-if="company.code === '71'">
              <product-configuration
                :editId="editId"
                :lineproductId="lineproductId"
                :currentCell="currentcell"
                class="q-pa-sm"
                :isCodeBuilder="isCodeBuilder"
                :drawingItemDetails="drawingItemDetails"
                @closeProdConfDialog="technicalData = false"
                @saveDrawingReqDetails="saveDrawingReqDetails"
              />

            <div class="row q-ml-md q-gutter-x-sm">
                <q-input
                  dense
                  min="0"
                  outlined
                  type="number"
                  label="Unit Price"
                  class="col-2"
                  :prefix="currency"
                  v-model.number="formdata['price']"
                  :rules="[val => val >= 0 || 'Value should be positive']"
                  @input="handlePriceInput"
                />

              <div v-if="formdata['multiplyval']" class="col-1">
                  <q-input
                    dense
                    v-model.number="formdata['multiplier']"
                    type="number"
                    min="0"
                    label="Multiplier"
                    outlined
                  />
                  </div>
                  <q-input
                    dense
                    disable
                    outlined
                    :label="`Unit ${currentcell.title} Price`"
                    type="number"
                    class="col-3"
                    :prefix="currency"
                    v-model.number="formdata['unitPrice']"
                  />
                  <q-input
                    dense
                    v-model.number="formdata['qty']"
                    type="number"
                    label="Qty"
                    min="1"
                    :rules="[val => val >= 1 || 'Value should be positive']"
                    class="col-1"
                    outlined
                  />
                  <q-input
                    dense
                    disable
                    outlined
                    :label="`Final ${currentcell.title} Price`"
                    type="number"
                    class="col-3"
                    :prefix="currency"
                    v-model="getfinalprice"
                  />
              </div>
            </div>
            </q-form>
          </q-card-section>
          <q-card-actions class="row q-mt-none">
            <q-btn
              dense
              size="xs"
              class="col-6"
              label="Save"
              color="primary"
              @click="saveDellData('save')"
            />
            <q-checkbox v-model="formdata['multiplyval']" label="Multiplier" />
            <q-checkbox v-model="formdata['outSource']" label="Out Sourced" />
          </q-card-actions>
        </q-card>
      </q-dialog>

      <q-dialog v-model="showseries" persistent>
        <q-card dense class="q-py-none" style="max-width: 1500px; width: 300px">
          <q-card-section class="row q-py-none">
            <div class="text-h6">Series</div>
            <q-space />
            <q-btn
              dense
              flat
              round
              size="xs"
              icon="close"
              v-close-popup
              class="float-right q-mt-xs q-pa-sm"
            >
              <q-tooltip
                max-height="100%"
                max-width="40%"
                content-class="bg-white text-black"
                >Close</q-tooltip
              >
            </q-btn>
          </q-card-section>

          <q-card-section class="q-pt-none q-py-none">
            <div class="col-12">
              <q-select
                v-model="series"
                dense
                outlined
                use-input
                clearable
                placeholder="Series"
                input-debounce="0"
                option-value="id"
                option-label="name"
                :options="serieslist"
                @filter="filterseries"
              >
                <template v-slot:no-option>
                  <q-item>
                    <q-item-section class="text-grey">
                      No results
                    </q-item-section>
                  </q-item>
                </template>
              </q-select>
              <!-- hide-dropdown-icon -->
            </div>
          </q-card-section>

          <q-card-actions align="center">
            <q-btn
              size="xs"
              class="col-12"
              label="Select"
              color="primary"
              v-close-popup
              @click="showitemcode(series)"
            />
          </q-card-actions>
        </q-card>
      </q-dialog>
      <q-dialog v-model="code" persistent>
        <q-card dense class="q-py-none" style="max-width: 1500px; width: 300px">
          <q-card-section class="row q-py-none">
            <div class="text-h6">Item Code</div>
            <q-space />
            <q-btn
              dense
              flat
              round
              size="xs"
              icon="close"
              v-close-popup
              class="float-right q-mt-xs q-pa-sm"
              @click="saveItemCode('close')"
            >
              <q-tooltip
                max-height="100%"
                max-width="40%"
                content-class="bg-white text-black"
              >
                Close
              </q-tooltip>
            </q-btn>
          </q-card-section>

          <q-card-section class="q-pt-none q-py-none">
            <div class="col-12">
              <q-select
                dense
                outlined
                use-input
                clearable
                option-value="id"
                input-debounce="0"
                v-model="itemCode"
                option-label="code"
                :options="listOfItemcode"
                placeholder="Series-Size-Rating-Body-Trim-Operator"
                @filter="filterItemCode"
                @input="cellItemcode['val']=itemCode ? itemCode.id : null"
              />
            </div>
          </q-card-section>

          <q-card-actions align="center">
            <q-btn
              size="xs"
              class="col-3"
              label="Back"
              color="accent"
              v-close-popup
              @click="showseries = true"
            />
            <q-btn
              size="xs"
              class="col-3"
              label="Select"
              color="primary"
              :loading="disableitemcodesave"
              :disable="disableitemcodesave"
              @click="selectItemCode"
              v-close-popup
            />
            <q-btn
              size="xs"
              class="col-5"
              color="primary"
              label="Show tech data"
              :loading="disableitemcodesave"
              :disable="disableitemcodesave"
              @click="showTechData"
            />
          </q-card-actions>
        </q-card>
      </q-dialog>
      <q-dialog v-if="actsizing" v-model="actsizing" persistent>
        <q-card dense class="q-py-none" style="max-width: 1500px; width: 1200px"
          ><q-form ref="sizingdata">
            <q-card-section class="row q-py-none">
              <div class="text-h6">Act Sizing</div>
              <div class="q-xl-sm">
                <div class="q-gutter-sm">
                  <q-radio
                    :val="2"
                    keep-color
                    label="SI"
                    color="cyan"
                    v-model="actsizingdata['makeId']"
                    @input="setActSizingDefaultData"
                  />
                  <q-radio
                    :val="3"
                    keep-color
                    color="cyan"
                    label="Imperial"
                    v-model="actsizingdata['makeId']"
                    @input="setActSizingDefaultData"
                  />
                </div>
              </div>
              <q-space />
              <q-btn
                dense
                flat
                round
                size="xs"
                icon="close"
                v-close-popup
                class="float-right"
                @click="saveSizing('close')"
              >
                <q-tooltip
                  max-height="100%"
                  max-width="40%"
                  content-class="bg-white text-black"
                  >Close</q-tooltip
                >
              </q-btn>
              <q-separator class="q-mt-md" />
            </q-card-section>
            <q-card-section>
              <div class="q-gutter-sm">
                <div class="row">
                  <div class="col-6 q-gutter-sm">
                    <div class="row">
                      <span class="col-4"> Actuator </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="name"
                        dense
                        outlined
                        class="col-8"
                        :rules="[rules.required]"
                        v-model.number="actsizingdata['actuator']"
                        :options="ListOfActuator"
                        use-input
                        @filter="filterActuator"
                        @input="actsizingdata['series'] = null"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                      </q-select>
                    </div>
                    <!-- v-if="actsizingdata['makeId'] === 3" // Imperial -->
                    <div class="row">
                      <span class="col-4"> Series </span>
                      <q-select
                        dense
                        outlined
                        use-input
                        options-dense
                        class="col-8"
                        option-value="id"
                        option-label="value"
                        :rules="[rules.required]"
                        v-model="actsizingdata['series']"
                        :options="actSizingSeriesOptions"
                        @filter="filterActuatorType"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                      </q-select>
                    </div>
                    <div class="row" v-if="actsizingdata['actuator'] && actsizingdata['actuator'].id === 185">
                      <span class="col-4"> Number of Phases </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="value"
                        class="col-8"
                        outlined
                        dense
                        v-model="actsizingdata['nosphaz']"
                        :options="ListOfPhase"
                        :rules="[rules.required]"
                        @filter="filterPhases"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                      </q-select>
                    </div>
                    <div class="row" v-else>
                      <span class="col-4"> Fail safe Condition </span>
                      <q-select
                        options-dense
                        dense
                        outlined
                        option-value="id"
                        option-label="value"
                        class="col-8"
                        :rules="[rules.required]"
                        v-model.number="actsizingdata['Shaft']"
                        :options="
                          Number(
                            this.actsizingdata.actuatorType
                              ? this.actsizingdata.actuatorType.id
                              : 207
                          ) === 207
                            ? [
                                {
                                  id: 165,
                                  value: 'S : Stayput'
                                }
                              ]
                            : [
                                {
                                  id: 163,
                                  value: 'C: Fail Close'
                                },
                                {
                                  id: 164,
                                  value: 'Z : Fail Open'
                                }
                              ]
                        "
                        use-input
                        @filter="filterFailSafe"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                      </q-select>
                    </div>
                    <div class="row">
                      <span class="col-4"> Special </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="value"
                        class="col-8"
                        outlined
                        dense
                        v-model="actsizingdata['special']"
                        :options="ListOfSpecial"
                        @filter="filterSpecial"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                        <template v-slot:option="scope">
                          <q-item
                            v-bind="scope.itemProps"
                            v-on="scope.itemEvents"
                          >
                            <q-item-section>
                              <q-item-label v-html="scope.opt.label" />
                              <q-item-label caption>{{ scope.opt.value }} :- {{ scope.opt.code }}</q-item-label>
                            </q-item-section>
                          </q-item>
                       </template>
                      </q-select>
                    </div>
                    <div class="row">
                      <span class="col-4"> Actuator Mounting </span>
                      <q-select
                        options-dense
                        dense
                        outlined
                        option-value="id"
                        option-label="value"
                        class="col-8"
                        v-model="actsizingdata['actMounting']"
                        :options="ListOfActuatorMounting"
                        :rules="[rules.required]"
                        use-input
                        @filter="filterActuatorMounting"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                      </q-select>
                    </div>
                    <div class="row">
                      <span class="col-4"> Valve Type </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="name"
                        class="col-8"
                        clearable
                        :disable="true"
                        dense
                        outlined
                        v-model="actsizingdata['valveType']"
                        :options="ListOfProductName"
                        use-input
                        @filter="filterProductName"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                      </q-select>
                    </div>
                    <div class="row">
                      <span class="col-4">
                        Differential Pressure ({{
                          getmeasureunit[actsizingdata.makeId === 2 ? 'Metric' : storedummydata.measureunit].PR
                        }})
                      </span>
                      <q-input
                        dense
                        outlined
                        class="col-8"
                        type="number"
                        :rules="[rules.required]"
                        placeholder="Type and hit enter"
                        v-model.number="actsizingdata['diffPr']"
                        @keyup.enter="actuatorTorqueValue"
                      />
                      <!-- filterDiffPressure -->
                    </div>
                    <div class="row">
                      <span class="col-4">
                        Valve Torque ({{
                          getmeasureunit[actsizingdata.makeId === 2 ? 'Metric' : storedummydata.measureunit].Torq
                        }})
                      </span>
                      <q-input
                        class="col-8"
                        type="number"
                        outlined
                        dense
                        :rules="[rules.required]"
                        v-model.number="actsizingdata['vtorque']"
                      />
                    </div>
                  </div>
                  <q-space />
                  <div class="col-6 q-gutter-sm">
                    <div class="row" v-if="actsizingdata['actuator'] && actsizingdata['actuator'].id === 185">
                      <span class="col-4"> Supply Voltage </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="value"
                        class="col-8"
                        outlined
                        dense
                        v-model="actsizingdata['supplyvoltage']"
                        :options="ListOfSupplyVoltage"
                        :rules="[rules.required]"
                        @filter="filterSupplyVoltage"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                        <template v-slot:option="scope">
                          <q-item
                            v-bind="scope.itemProps"
                            v-on="scope.itemEvents"
                          >
                            <q-item-section>
                              <q-item-label v-html="scope.opt.label" />
                              <q-item-label caption>{{ scope.opt.value }} :- {{ scope.opt.code }}</q-item-label>
                            </q-item-section>
                          </q-item>
                       </template>
                      </q-select>
                    </div>
                    <div class="row" v-else>
                      <span class="col-4"> Actuator Type </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="value"
                        class="col-8"
                        outlined
                        dense
                        v-model="actsizingdata['actuatorType']"
                        :options="ListOfActuatorType"
                        :rules="[rules.required]"
                        use-input
                        @filter="filterActuatorType"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                      </q-select>
                    </div>
                    <div class="row" v-if="actsizingdata['actuator'] && actsizingdata['actuator'].id === 185">
                      <span class="col-4"> Hertz </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="value"
                        class="col-8"
                        outlined
                        dense
                        v-model="actsizingdata['hertz']"
                        :options="ListOfHertz"
                        :rules="[rules.required]"
                        @filter="filterHertz"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                        <template v-slot:option="scope">
                          <q-item
                            v-bind="scope.itemProps"
                            v-on="scope.itemEvents"
                          >
                            <q-item-section>
                              <q-item-label v-html="scope.opt.label" />
                              <q-item-label caption>{{ scope.opt.value }} :- {{ scope.opt.code }}</q-item-label>
                            </q-item-section>
                          </q-item>
                       </template>
                      </q-select>
                    </div>
                    <div class="row" v-else>
                      <span class="col-4">
                        Supply Pressure ({{
                          getmeasureunit[actsizingdata.makeId === 2 ? 'Metric' : storedummydata.measureunit].PR
                        }})
                      </span>
                      <q-select
                        label="min"
                        options-dense
                        option-value="code"
                        option-label="value"
                        dense
                        outlined
                        use-input
                        v-model="actsizingdata['SupplyPRMin']"
                        :options="ListOfSupplyPressure"
                        :rules="[rules.required]"
                        class="col-3"
                        @filter="filterSupplyPressure"
                      />
                      <q-space />
                      <q-select
                        label="max"
                        options-dense
                        option-value="id"
                        option-label="value"
                        use-input
                        dense
                        outlined
                        v-model="actsizingdata['SupplyPRMax']"
                        :options="ListOfSupplyPressure"
                        class="col-3"
                        @filter="filterSupplyPressure"
                      />
                    </div>
                    <div class="row" v-if="actsizingdata['actuator'] && actsizingdata['actuator'].id === 185">
                      <span class="col-4"> Enclosure Type </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="value"
                        class="col-8"
                        outlined
                        dense
                        use-input
                        v-model="actsizingdata['enclosure']"
                        :options="ListOfEnclosure"
                        :rules="[rules.required]"
                        @filter="filterEnclosure"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                        <template v-slot:option="scope">
                          <q-item
                            v-bind="scope.itemProps"
                            v-on="scope.itemEvents"
                          >
                            <q-item-section>
                              <q-item-label v-html="scope.opt.label" />
                              <q-item-label caption>{{ scope.opt.value }} :- {{ scope.opt.code }}</q-item-label>
                            </q-item-section>
                          </q-item>
                       </template>
                      </q-select>
                    </div>
                       <div class="row" >
                      <span class="col-4"> Insert </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="value"
                        class="col-8"
                        outlined
                        use-input
                        dense
                        v-model="actsizingdata['insert']"
                        :options="ListOfInsert"
                        @filter="filterInsert"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                        <template v-slot:option="scope">
                          <q-item
                            v-bind="scope.itemProps"
                            v-on="scope.itemEvents"
                          >
                            <q-item-section>
                              <q-item-label v-html="scope.opt.label" />
                              <q-item-label caption>{{ scope.opt.value }} :- {{ scope.opt.code }}</q-item-label>
                            </q-item-section>
                          </q-item>
                       </template>
                      </q-select>
                    </div>
                    <div class="row" v-if="actsizingdata['actuator'] && actsizingdata['actuator'].id === 193">
                      <span class="col-4"> Manual Override </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="value"
                        class="col-8"
                        outlined
                        dense
                        use-input
                        v-model="actsizingdata['manualOverride']"
                        :options="ListOfMOR"
                        :rules="[rules.required]"
                        @filter="filterMOR"
                      >
                      <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                        <template v-slot:option="scope">
                          <q-item
                            v-bind="scope.itemProps"
                            v-on="scope.itemEvents"
                          >
                            <q-item-section>
                              <q-item-label v-html="scope.opt.label" />
                              <q-item-label caption>{{ scope.opt.value }} :- {{ scope.opt.code }}</q-item-label>
                            </q-item-section>
                          </q-item>
                       </template>
                      </q-select>
                    </div>
                    <div class="row" v-else>
                      <span class="col-4"> Manual Override </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="value"
                        class="col-8"
                        use-input
                        dense
                        outlined
                        v-model="actsizingdata['manualOverride']"
                        :options="ListOfManualOverride"
                        @filter="filterManualOverride"
                      >
                        <template v-slot:no-option>
                          <q-item>
                            <q-item-section class="text-grey">
                              No results
                            </q-item-section>
                          </q-item>
                        </template>
                      </q-select>
                    </div>
                    <div class="row">
                      <span class="col-4"> Media</span>
                      <q-input
                        class="col-8"
                        :rules="[rules.required]"
                        type="text"
                        dense
                        outlined
                        v-model="actsizingdata['Media']"
                      />
                    </div>
                    <div class="row">
                      <span class="col-4"> Media Type </span>
                      <q-select
                        class="col-8"
                        options-dense
                        option-value="id"
                        option-label="value"
                        use-input
                        dense
                        outlined
                        v-model="actsizingdata['mediaType']"
                        :rules="[rules.required]"
                        :options="ListOfMediaType"
                        @filter="filterMediaType"
                      />
                    </div>
                    <div class="row">
                      <span class="col-4"> Safety Factor </span>
                      <q-select
                        options-dense
                        option-value="id"
                        option-label="value"
                        use-input
                        v-model="actsizingdata['SafetyFactor']"
                        :options="ListOfSafetyFactor"
                        dense
                        class="col-8"
                        outlined
                        @filter="filterSafetyFactor"
                        :rules="[rules.required]"
                      />
                    </div>
                  </div>
                </div>
                <div class="row">
                  <span class="col-2"> Internal Note </span>
                  <q-input
                    class="col-10"
                    dense
                    outlined
                    v-model="actsizingdata['note']"
                  />
                </div>
                <div class="col-12">
                  <q-btn
                    dense
                    size="xs"
                    color="primary"
                    class="full-width"
                    label="Search Sizes"
                    @click="getActuator"
                  />
                </div>
                <div class="row">
                  <span class="col-3"> FG Code </span>
                  <q-select
                    dense
                    outlined
                    use-input
                    clearable
                    class="col-9"
                    v-model="actsizingdata['fgcode']"
                    :options="actuatorItemCodeSelectOptions"
                    @filter="filterFgCode"
                    @input="getDescriptionOnChangeOfFgCode"
                  ></q-select>
                </div>
                <div class="row">
                  <span class="col-3"> FG Description </span>
                  <q-input
                    class="col-9"
                    outlined
                    dense
                    v-model="actsizingdata['fgDesc']"
                  />
                </div>
              </div>
            </q-card-section>
            <q-card-actions align="center">
              <q-btn
                v-show="
                  actsizingdata['fgcode'] && actsizingdata['fgcode'].length
                "
                dense
                size="xs"
                class="col-12"
                v-close-popup
                label="Select"
                color="primary"
                @click="saveDataInXl"
              />
            </q-card-actions>
          </q-form>
        </q-card>
      </q-dialog>
      <q-dialog persistent v-model="technicalData">
        <div style="width: 900px; max-width: 80vw">
          <technical-data-sheet
            v-if="storedummydata.columns"
            :column="currentcell"
            :current-column="curColumn"
            :model="
              companydata.code === '08_DELVAL_USA'
                ? formdata['itemcode'] || (cellItemcode && cellItemcode['val'])
                : formdata['model']
            "
            :is-show-tech-data="isShowTechData"
            :make="storedummydata.columns.makeId"
            :tech-query-variable="techQueryVariable"
            :productversion="currentProductVersion"
            :technical-data="storedummydata && storedummydata.technicaldata"
            @close-popup="technicalData = false"
            @tech-data-updated="techDataUpdateEvent"
          />
        </div>
      </q-dialog>

      <!-- Company Group Dialog -->
      <q-dialog
        persistent
        full-width
        full-height
        v-model="addNewOptionDialog"
        transition-show="slide-up"
        transition-hide="slide-down"
      >
        <q-card>
          <q-bar class="bg-primary text-white">
            <div class="text-h6">
              <q-icon class="q-mr-sm" size="xs" name="domain" />
              Version
            </div>
            <q-space />
            <q-btn dense flat size="xs" icon="close" v-close-popup @click="formData = null">
              <q-tooltip max-height="30%" max-width="25%">
                <div class="tooltip-content">
                  Close
                </div>
              </q-tooltip>
            </q-btn>
          </q-bar>

          <q-card-section>
            <version-form
              @closeDialog="addNewOptionDialog = false"
              @recordSaved="handleProductSavedEvent"
            />
          </q-card-section>
        </q-card>
      </q-dialog>
    </div>
  </div>
</template>

<script>
// eslint-disable-next-line
/* eslint-disable */
import Vue from 'vue'
import gql from 'graphql-tag'
import Component from 'vue-class-component'
import { Watch } from 'vue-property-decorator'
import { mapGetters, mapActions, mapMutations, mapState } from 'vuex'

import { queries, mutations } from '@/graphql'
import { masterTechData } from '@/graphql/proposal.gql'
import jspreadsheet, { jexcel } from '@/packages/jspreadsheet'
import {
  Message,
  CommonUtility,
  ProposalUtility,
  DataConstructor
} from '@/mixins'
import {
  TemplateMutation,
  NonStandardMutation,
  NonStandardSubscription,
  NonStandardPagination
} from '@/services'

var itemcodelist = []
var filteredlist = []
var nsDependentopt = []
var nsOptions = []

@Component({
  name: 'jexcel',
  mixins: [
    Message,
    CommonUtility,
    ProposalUtility,
    DataConstructor,
    TemplateMutation,
    NonStandardMutation,
    NonStandardSubscription,
    NonStandardPagination
  ],
  props: [
    'newsheet',
    'destroy',
    'saveexceldata',
    'newsheetcomment',
    'validatePageBeforeSave'
  ],
  computed: {
    ...mapState('proposal', ['proposaldata']),
    ...mapGetters('proposal', [
      'getProposaldata',
      'getmeasureunit',
      'getcompanytax',
      'getPriceList'
    ]),
    ...mapGetters('common', [
      'salesRegion',
      'drawerstate',
      'qSpinner',
      'company'
    ]),
    ...mapGetters('auth', ['user']),
    ...mapGetters('userconfig', ['getUsercolumn'])
  },
  methods: {
    ...mapActions('proposal', [
      'updateproposalRowdata',
      'updateValidation',
      'updatecompanytax',
      'addPriceList',
      'replacePriceList'
    ]),
    ...mapMutations('common', ['toggleSpinner', 'setOfferId'])
  },
  components: {
    'version-form': () => import('@/pages/configurations/product/versionForm.vue'),
    'technical-data-sheet': () => import('components/proposal/technicaldatatable.vue'),
    'product-configuration': () => import('components/proposal/productConf.vue')
  }
})
export default class ProposalExcel extends Vue {
  data() {
    return {
      valepopsetting: false,
      rowpaste: true,
      lineproductId: '',
      productMetaData: {
        productId: null
      },
      isShow: false,
      editId: null,
      isEdit: false,
      tempCopySourceData: [],
      roundCount: 0,
      decimalSeparator: 0,
      thousandSeparator: 0,
      currentProduct: null,
      currentSeries: null,
      techQueryVariable: {},
      getTorqueListForSeries: [],
      isShowTechData: false,
      oldSeriesId: null,
      notif: null,
      notifPercentage: 0,
      rules: {
        required: value => !!value || this.$t('validation.required')
      },
      actuatorItemcodeList: [],
      actuatorItemCodeSelectOptions: [],
      currentsheetData: null,
      currentPriceList: null,
      taxpriceval: {},
      discountprice: null,
      sheetinstace: null,
      rowvalveprice: null,
      taxes: false,
      offerTax: {},
      nonStandardList: nsOptions,
      nsDependendlist: nsDependentopt,
      nsDependentvalue: null,
      proposalNS: [],
      newNSItems: [],
      itemcodeexceldata: [],
      disableitemcodesave: false,
      selectedtab: 'primary',
      alphabetlist: [],
      modellist: [],
      actuatorprice: null,
      copiedcells: null,
      color: '',
      series: null,
      serieslist: [],
      showseries: false,
      technicalData: false,
      actsizing: false,
      actsizingdata: {
        fgcode: null,
        makeId: 2,
        series: '21',
        rowNumber: 0
      },
      cellItemcode: {
        y: 0,
        val: ''
      },
      companydata: null,
      comId: 0,
      orignalcolumns: null,
      itemCode: null,
      formdata: {
        desc: '',
        price: 1,
        multiplier: 1,
        qty: 1,
        itemcode: null,
        multiplyval: true
      },
      isCodeBuilder:false,
      comment: false,
      currentProductVersion: {},
      currentcell: { val: null, pos: [], title: '' },
      make: '',
      sheetype: null,
      currency: '',
      templatedata: null,
      curunitcost: 0,
      priceindexlist: [],
      ListOfProductName: [],
      ListOfMediaType: [],
      ListOfDiffPressure: [],
      ListOfSafetyFactor: [],
      ListOfSupplyPressure: [],
      ListOfHertz: [],
      ListOfSupplyVoltage: [],
      ListOfFailSafeCondition: [],
      ListOfPhase: [],
      ListOfSpecial: [],
      ListOfInsert: [],
      ElecActData: [],
      ListOfEnclosure: [],
      ListOfMOR: [],
      ListOfActuator: [],
      ListOfActuatorType: [],
      ListOfActuatorMounting: [],
      ListOfManualOverride: [],
      nsMdl: false,
      nsText: null,
      nsColumnNameTarget: {},
      secondDialog: false,
      model: null,
      options: [],
      excel: [],
      dummydata: [],
      storedummydata: {},
      dataindex: 1,
      dataindex1: ' ',
      dropdownlist: {},
      setrow: true,
      code: false,
      previtemcodes: {},
      trimDataCode: [],
      listOfItemcode: [],
      deletedcoumn: { metaData: {}, technicaldata: {}, sizingdata: {} },
      deletedRowdata: { metaData: {}, technicaldata: {}, sizingdata: {} },
      usaNestedHeader: [
        { title: 'General', colspan: '8' },
        { title: 'Valve', colspan: '19' },
        { title: 'Act Sizing', colspan: '5' },
        { title: 'Valve LP', colspan: '1' },
        { title: 'Operator', colspan: '5' },
        { title: 'Accessories', colspan: '8' },
        { title: 'Additional Charges', colspan: '2' },
        { title: 'Final Price', colspan: '4' },
        { title: 'Delivery', colspan: '2' },
        { title: 'Note', colspan: '1' }
      ],
      curColumn: null,
      addNewOptionDialog: false,
      currentEditData: null
    }
  }

  get getfinalprice() {
    if (_.isEmpty(this.formdata)) {
      this.formdata = {
        desc: '',
        price: null,
        qty: 1,
        unitPrice: null,
        multiplier: 1,
        itemcode: null,
        multiplyval: true
      }
    }

    const price = !this.formdata?.['price']
      ? 0
      : this.formdata['price']
    const multiplier = !this.formdata['multiplier']
      ? 1
      : this.formdata['multiplier']
    const qty = !this.formdata['qty']
      ? 1
      : this.formdata['qty']

    let unitPrice = 0
    if (this.company.code === '71'){
      unitPrice = parseFloat(price / multiplier).toFixed(this.roundCount)
      } else {
      unitPrice = parseFloat(price * multiplier).toFixed(this.roundCount)
      }
    this.formdata['unitPrice'] = unitPrice
    let finalPrice = parseFloat(unitPrice * qty).toFixed(this.roundCount)
    return finalPrice
  }

  get jExcelOptions() {
    return {
      data: [],
      newcol: [],
      search: false,
      wordWrap: true,
      lazyLoading: false,
      sheetName: 'Delval USA',
      minDimensions: [15, 50],
      defaultColWidth: 65,
      columnDrag: true, // NOTE Column drag option
      allowComments: true,
      allowInsertColumn: this.company?.code === '03_DELVAL',
      allowundo: true,
      autoIncrement: false,
      tableOverflow: true,
      loadingSpin: true,
      allowExport: true,
      tableWidth: '88.5vw', // '93vw',
      tableHeight: '60vh',
      filters: true,
      meta: {},
      // csvHeaders: true,
      columnSorting: false,
      freezeColumns: 0,
      columns: [],
      nestedHeaders: [],
      contextMenu: function(obj, x, y, e) {
        var items = []

        if (y == null) {
          // Insert a new column
          if (obj.options.allowInsertColumn === true) {
            items.push({
              title: obj.options.text.insertANewColumnBefore,
              onclick: function() {
                obj.insertColumn(1, parseInt(x), 1)
              }
            })
          }

          if (obj.options.allowInsertColumn === true) {
            items.push({
              title: obj.options.text.insertANewColumnAfter,
              onclick: function() {
                obj.insertColumn(1, parseInt(x), 0)
              }
            })
          }

          // Delete a column
          if (obj.options.allowDeleteColumn === true) {
            items.push({
              title: obj.options.text.deleteSelectedColumns,
              onclick: function() {
                if (confirm('Are you sure to delete the selected columns?')) {
                  obj.deleteColumn(
                    obj.getSelectedColumns().length
                      ? undefined
                      : parseInt(_.cloneDeep(x))
                  )
                }
              }
            })
          }

          // Rename column
          if (obj.options.allowRenameColumn === true) {
            items.push({
              title: obj.options.text.renameThisColumn,
              onclick: function() {
                obj.setHeader(x)
              }
            })
          }
          // undo
          if (obj.historyIndex >= 0) {
            items.push({
              title: obj.options.text.undo,
              onclick: function() {
                obj.undo()
              }
            })
          }
          // redo
          if (obj.historyIndex < obj.history.length - 1) {
            items.push({
              title: obj.options.text.redo,
              shortcut: 'Ctrl + Y',
              onclick: function() {
                obj.redo()
              }
            })
          }

          // Sorting
          if (obj.options.columnSorting === true) {
            // Line
            items.push({ type: 'line' })

            items.push({
              title: obj.options.text.orderAscending,
              onclick: function() {
                obj.orderBy(x, 0)
              }
            })
            items.push({
              title: obj.options.text.orderDescending,
              onclick: function() {
                obj.orderBy(x, 1)
              }
            })
          }
        } else {
          // Insert new row
          if (obj.options.allowInsertRow === true) {
            items.push({
              title: obj.options.text.insertANewRowBefore,
              onclick: function() {
                obj.insertRow(1, parseInt(y), 1)
              }
            })

            items.push({
              title: obj.options.text.insertANewRowAfter,
              onclick: function() {
                obj.insertRow(1, parseInt(y))
              }
            })
          }

          if (obj.options.allowDeleteRow === true) {
            items.push({
              title: obj.options.text.deleteSelectedRows,
              onclick: function() {
                if (
                  confirm(
                    'Additional cell data might not  be recovered. Are you sure to delete the selected row.?'
                  )
                ) {
                  obj.deleteRow(
                    obj.getSelectedRows().length ? undefined : parseInt(y)
                  )
                }
              }
            })
          }
          // undo
          if (obj.historyIndex >= 0) {
            items.push({
              title: obj.options.text.undo,
              onclick: function() {
                obj.undo()
              }
            })
          }
          // redo
          if (obj.historyIndex < obj.history.length - 1) {
            items.push({
              title: obj.options.text.redo,
              shortcut: 'Ctrl + Y',
              onclick: function() {
                obj.redo()
              }
            })
          }

          if (x) {
            if (obj.options.allowComments === true) {
              items.push({ type: 'line' })

              var title = obj.records[y][x]?.getAttribute('title') || ''

              items.push({
                title: title
                  ? obj.options.text.editComments
                  : obj.options.text.addComments,
                onclick: function() {
                  obj.setComments(
                    [x, y],
                    prompt(obj.options.text.comments, title)
                  )
                }
              })

              if (title) {
                items.push({
                  title: obj.options.text.clearComments,
                  onclick: function() {
                    obj.setComments([x, y], '')
                  }
                })
              }
            }
          }
        }

        // sort
        if (obj.options.columnSorting === true) {
          items.push({
            title: obj.options.text.orderAscending,
            onclick: function() {
              obj.orderBy(x, 0)
            }
          })
          items.push({
            title: obj.options.text.orderDescending,
            onclick: function() {
              obj.orderBy(x, 1)
            }
          })
        }
        // Line
        items.push({ type: 'line' })

        // Save
        if (obj.options.allowExport) {
          items.push({
            title: obj.options.text.saveAs,
            shortcut: 'Ctrl + S',
            onclick: function() {
              obj.download(true)
            }
          })
        }

        // filter
        if (obj.options.filter) {
          items.push({
            title: obj.options.text.filter,
            onclick: function() {
              // alert(obj.options.filter);
            }
          })
        }

        return items.filter(row => row.title)
      },
      onchange: this.changed,
      onbeforechange: this.beforeChange,
      ondeletecolumn: this.recalc,
      onchangeheader: this.changeheader,
      onbeforedeleterow: this.beforedeleterow,
      onundo: this.undochanges,
      onredo: this.redochange,
      oninsertcolumn: this.insertcol,
      // onbeforepaste: this.beforepaste,
      onchangemeta: this.changemeta,
      onpaste: this.paste,
      onbeforeinsertrow: this.onbeforeinsertrow,
      oninsertrow: this.onInsertRow,
      ondeleterow: this.deleterow,
      onbeforepaste: this.beforepaste,
      onmouseover: this.mousehover,
      onbeforedeletecolumn: this.beforedelcol,
      oncopy: this.copycell,
      ondeletecell: this.deletecell,
      oneditionstart: this.editionstart,
      onselection: this.onSelection,
      // onfocus: this.exceRowFocus,
      // onmouseover: this.focus,
      // onblur: this.blur,
      // oneditionend: this.editionend,
      prompt: this.promptMethod,
      onCopyRecords: this.recordCopy
    }
  }

  get refoptions() {
    let options = null
    debugger
    if (this.getProposaldata?.proposalRefId) {
      const column = this.getProposaldata.proposalRefId?.columns

      if (column !== undefined) {
        options = {
          ..._.cloneDeep(this.jExcelOptions),
          sheetName: 'Delval India ',
          columns: column?.columns
        }
        column['sheetName'] = 'Delval India '
      }
    }
    return options
  }

  get interCoComId() {
    return this.getProposaldata?.proposalRefId?.com?.id
  }


  get actSizingSeriesOptions() {
    if (this.actsizingdata['actuator']?.id === 185) {
      this.actsizingdata['series'] = '2E'
    }
    return (_.includes([184, 8],this.actsizingdata['actuator']?.id))
      ? ['21', '2Z']
      : (this.actsizingdata['actuator']?.id === 185) ? ['2E'] : ['25']
  }

  get getWidth() {
    return this.drawerstate
      ? window.screen.width - 320
      : window.screen.width - 200
  }

  get drawingItemDetails() {
    return this.productMetaData
  }

  set drawingItemDetails (value) {
    this.productMetaData = value
  }

  @Watch('company', { immediate: true, deep: true })
  handlecompanychanges(val) {
    this.companydata = val || this.user.company
  }

  @Watch('drawerstate', { deep: true })
  handlejexcel(val) {
    const sheet = this.$refs.spreadsheet.jexcel
    const jexcelwidth = !val.drawer ? '90vw' : '83.5vw'
    this.$set(sheet.options, 'tableWidth', jexcelwidth)
    // if (document.querySelector('.jexcel_content') && this.getWidth) {
    //   document.querySelector('.jexcel_content').style.width = `${this.getWidth}px`
    // }
  }

  @Watch('destroy', { deep: true })
  handledestroy(val) {
    if (val === true) {
      this.savedata()
    }
  }

  @Watch('getfinalprice', { deep: true })
  handlefinalprice(newVal) {
    this.formdata['finalprice'] = newVal
  }

  created() {
    this.$emit('excel-instance', this)
    this.storedummydata = _.cloneDeep(this.getProposaldata)
    this.roundCount = (
      this.companydata?.roundCount || 2
    )
    this.decimalSeparator = this.companydata?.decimalSeparator || '.'
    this.thousandSeparator = this.companydata?.thousandSeparator || ''
  }

  mounted() {
    this.toggleSpinner(true)
    if (this.companydata.code === '08_DELVAL_USA') {
      this.jExcelOptions.nestedHeaders = this.usaNestedHeader
    }

    this.$set(
      this.actsizingdata,
      'makeId',
      (this.storedummydata?.measureunit === 'Metric') ? 2 : 3
    )
    this.$nextTick(() => {
      this.setActSizingDefaultData()
    })
    let offerSheetHasLength = false

    if (this.storedummydata?.offerRowData?.length) {
      offerSheetHasLength = true
    }
    this.currency = this.storedummydata?.currencyId?.symbol || ''

    this.$worker.postMessage([
      JSON.stringify([
        this.storedummydata,
        this.currency,
        offerSheetHasLength,
        this.companydata
      ]),
      'set-column-for-sheet'
    ])
    this.$worker.onmessage = e => {
      if (_.last(e.data) === 'set-column-for-sheet') {
        const data = _.first(e.data)
        const columns = data?.columns

        _.forEach(data?.filterIndex, x => {
          if (columns[x]) {
            columns[x]['filter'] = this.dropdownFilter
            columns[x]['sourceorignal'] = columns[x]?.source
          }
        })

        setTimeout(() => {
          this.constructexcelInstance(columns)
        }, 100)
      }
    }
  }

  promptMethod(message, oldVal, column) {
    return new Promise((resolve, reject) => {
      if (column?.defaultcol) {
        resolve(oldVal)
        this.showInfoMsg({ message: 'You can\'t edit default column.!' })
        return
      }

      this.$q.dialog({
        message,
        prompt: {
          model: oldVal || '',
          type: 'text' // optional
        },
        cancel: true,
        persistent: true
      }).onOk(newVal => {
        resolve(newVal)
      }).onCancel(() => {
        resolve(oldVal)
      }).onDismiss(() => {
        resolve(oldVal)
      })
    })
  }

  recordCopy(instance, records, sourceIndex) {
    const [x1, y1] = sourceIndex
    const metaData = instance.jexcel.getMeta()

    _.forEach(records, (record) => {
      const { x, y, col, row, newValue, oldValue } = record
      const sourceCellName = instance.jexcel.headers[x]?.innerText
      let meta = metaData[sourceCellName]?.[y1]
      const cellName = instance.jexcel.headers[col]?.innerText

      if (!_.includes(['ItemCode', 'fgCodeType'], cellName)) {
        meta = meta || { desc: null, price: newValue }
        instance.jexcel.setMeta(cellName, String(row), meta)

        if (meta.desc) {
          instance.jexcel.setComments([x, y], meta.desc)
        }
      }
    })
  }

  beforeDestroy() {
    if (this.$refs.spreadsheet?.jexcel) {
      this.$refs.spreadsheet.jexcel.destroy(true)
    }
  }

  destroyed() {
    this.storedummydata = null
    this.$emit('excel-instance', null)
  }

  handlePriceInput() {
    this.$set(
      this.formdata,
      'price',
      parseFloat(this.formdata?.price || 0).toFixed(this.roundCount)
    )
  }

  // el, rowNumber, numOfRows, insertBefore
  onbeforeinsertrow(instance, rowNumber, numOfRows, insertBefore, isLastRow) {
    if (isLastRow !== true) {
      const tempMetaData = {}
      const metaData = instance.jexcel.getMeta()
      const metaKeys = _.keys(metaData)

      _.forEach(metaKeys, (key) => {
        let index = 0
        let fgIndex = 0
        let trimIndex = 0

        if (!tempMetaData[key]) {
          tempMetaData[key] = {}
        }

        if (key === 'fgCodeType') {
          tempMetaData[key] = metaData[key]
        }

        const metaKeys = _.keys(metaData[key])

        _.forEach(metaKeys, (rowIndex) => {
          if (key === 'ItemCode') {
            const [itemCodeKey, ind] = _.split(rowIndex, '_')

            if (_.includes(_.toLower(itemCodeKey), 'trim')) {
              trimIndex = Number(ind) + 1

              if (Number(ind) === rowNumber) {
                tempMetaData[key][`Trim_${ind}`] = []
                tempMetaData[key][`Trim_${trimIndex}`] = metaData[key][rowIndex]
              } else if (Number(ind) < rowNumber) {
                tempMetaData[key][rowIndex] = metaData[key][rowIndex]
              } else if (Number(ind) > rowNumber) {
                tempMetaData[key][`Trim_${trimIndex}`] = metaData[key][rowIndex]
              }
            }

            if (_.includes(_.toLower(itemCodeKey), 'fgcode')) {
              fgIndex = Number(ind) + 1

              if (Number(ind) === rowNumber) {
                tempMetaData[key][`fgcode_${ind}`] = []
                tempMetaData[key][`fgcode_${fgIndex}`] = metaData[key][rowIndex]
              } else if (Number(ind) < rowNumber) {
                tempMetaData[key][rowIndex] = metaData[key][rowIndex]
              } else if (Number(ind) > rowNumber) {
                tempMetaData[key][`fgcode_${fgIndex}`] = metaData[key][rowIndex]
              }
            }
          } if (key === 'fgCodeType') {
            // fgCode logic
          } else {
            index = Number(rowIndex) + 1

            if (Number(rowIndex) === rowNumber) {
              tempMetaData[key][rowIndex] = { desc: null, price: 0 }
              tempMetaData[key][index] = metaData[key][rowIndex]
            } else if (Number(rowIndex) < rowNumber) {
              tempMetaData[key][rowIndex] = metaData[key][rowIndex]
            } else if (Number(rowIndex) > rowNumber) {
              tempMetaData[key][index] = metaData[key][rowIndex]
            }
          }
        })
      })

      instance.jexcel.options.meta = tempMetaData
    }
  }

  onInsertRow(instance, rowNumber, numOfRows, rowRecords, insertBefore) {}

  onSelection(instance, x, y, value) {
    let jsonData = _.filter(instance.jexcel.getJson(), 'Product')
    const indexQty = _.findIndex(this.jExcelOptions.columns, 'qty')
    const indextotal = _.findIndex(this.jExcelOptions.columns, 'totalval')
    const productColumn = _.findIndex(this.jExcelOptions.columns, 'productcol')

    if (
      !jsonData?.length ||
      !instance.jexcel.getValueFromCoords(productColumn, y)
    ) {
      jsonData = instance.jexcel.getJson()
    }

    if (indextotal !== -1) {
      const totalval1 = _.map(
        jsonData,
        tval =>
          parseFloat(tval[this.jExcelOptions.columns[indextotal]?.title]) || 0
      )

      const sumdata = {
        qty: _.sum(
          _.map(
            jsonData,
            b =>
              parseFloat(b[instance.jexcel.options.columns[indexQty]?.title]) ||
              0
          )
        ),
        total: _.sum(totalval1)
      }

      this.$emit('sumdata', sumdata)
    }

    if (
      this.jExcelOptions.columns?.[Number(x)]?.id !== 'Series' &&
      this.jExcelOptions.columns?.[Number(x)]?.id !== 'Valve_Size'
    ) {
      const indexseries = _.findIndex(this.jExcelOptions.columns, [
        'id',
        'Series'
      ])
      const seriesval = instance.jexcel.getValueFromCoords(indexseries, y)
      const indexsize = _.findIndex(this.jExcelOptions.columns, [
        'sourceId.name',
        'SIZE'
      ])
      const sizeval = instance.jexcel.getValueFromCoords(indexsize, y)
      const indexqty = _.findIndex(this.jExcelOptions.columns, 'qty')
      const qtyval = instance.jexcel.getValueFromCoords(indexqty, y)

      this.$emit('series', {
        series: seriesval,
        size: sizeval,
        Qty: qtyval
      })
    }
  }

  getProductDependentColumData(id) {
    return this.$apollo.query({
      query: gql`
        query product($id: ID!) {
          product: productVersion(id: $id) {
            id
            cgst
            igst
            hsnCode
            shortDes
            sgstOrUgst
          }
        }
      `,
      variables: { id }
    }).then(({ data }) => (data?.product || {}))
  }

  setActSizingDefaultData() {
    let makeId = this.actsizingdata?.makeId
    if (!makeId) {
      makeId = (this.storedummydata?.measureunit === 'Metric') ? 2 : 3
    }

    const SupplyPRMin = (makeId === 2) // NOTE SI
      ? {
        id: 184,
        value: '4',
        code: '4',
        masterFieldId: 151,
        companyId: [1],
        data: { makeId: { id: 2, name: 'SI' } }
      } : {
        id: 194,
        value: '60',
        code: '4',
        masterFieldId: 151,
        companyId: [1],
        data: { makeId: { id: 3, name: 'Imperial' } }
      }
    const SupplyPRMax = (makeId === 2) // NOTE SI
      ? {
        id: 190,
        value: '8',
        code: '8',
        masterFieldId: 151,
        companyId: [1],
        data: { makeId: { id: 2, name: 'SI' } }
      } : {
        id: 200,
        value: '115',
        code: '8',
        masterFieldId: 151,
        companyId: [1],
        data: { makeId: { id: 3, name: 'Imperial' } }
      }

    let usaDefaultData
    let manualOverride

    if (this.companydata.code === '08_DELVAL_USA') {
      usaDefaultData = {
        Media: 'Air',
        mediaType: { id: 177, value: 'Gas', code: null, masterFieldId: 149, companyId: [1], data: { makeId: null } },
        SafetyFactor: { id: 202, value: '1.3', code: null, masterFieldId: 152, companyId: [1], data: { makeId: null } }
      }
    } else {
       manualOverride = {
        id: 212,
        value: 'No',
        code: null,
        masterFieldId: 155,
        companyId: [1],
        data: { makeId: null }
      }
    }

    this.actsizingdata = {
      // ...this.actsizingdata,
      makeId,
      diffPr: null,
      vtorque: null,
      fgDesc: null,
      fgcode: null,
      series: '21',
      ...usaDefaultData,
      SupplyPRMin,
      SupplyPRMax,
      actMounting: {
        id: 209,
        value: 'Parallel',
        code: null,
        masterFieldId: 154,
        companyId: [1],
        data: { makeId: null }
      }
    }
  }

  createItemCode (val, done) {
    done(val)
    this.formdata['itemcode'] = val

    if (!itemcodelist[val]) {
      itemcodelist[val] = val
    }
  }

  filterFgCode(val, update) {
    update(() => {
      const actuatorItemcodeList1 = _.filter(
        this.actuatorItemcodeList,
        (ic) => _.startsWith(ic, this.actsizingdata?.series)
      )

      if (val === '') {
        this.actuatorItemCodeSelectOptions = actuatorItemcodeList1
      } else {
        const needle = val.toUpperCase()
        this.actuatorItemCodeSelectOptions = _.filter(
          actuatorItemcodeList1,
          v => _.includes(v, needle)
        )
      }
    })
  }

  getDescriptionOnChangeOfFgCode(fgCode) {
    if (fgCode) {
    let series = fgCode.split('-')[0]
    let fgCode1 = fgCode.split('-').join('')
      let verId = (this.actsizingData?.actuator?.id === 184 ? 8 : (this.actsizingData?.actuator?.id === 193 ? 9 : this.actsizingData?.actuator?.id )) || null
      if (verId === null){
        if (_.includes(['21','2Z'], series)){
          verId = 8
        } else if (_.includes(['25'], series)) {
          verId = 8
        } else {
          verId = 10
        }
      }
      this.$apollo.query({
        query: queries.getFgCodeDescription,
        variables: {
          verId: String(verId),
          fgCode: String(fgCode1),
          makeId: String(this.storedummydata?.columns?.makeId?.id)
        }
      }).then(({ data }) => {
        console.log(data,'data')
        this.$set(this.actsizingdata, 'fgDesc', data?.description || this.actsizingdata['fgDesc'])
        this.saveSizing('update')
      })
    }
  }

  techDataUpdateEvent() {
    this.$emit('tech-data-updated')

    this.$nextTick(() => {
      this.code = false
      this.isShowTechData = false
    })
  }

  saveDrawingReqDetails (val) {
    // console.log('SaveDrawing', val)
    this.isShow = false
    this.formdata['productMeta'] = val
    // check for cell meta for changed values
    const cell = this.currentcell
    this.productMetaData = this.formdata['productMeta']
    this.formdata['descLg'] = val.longDes
    this.formdata['desc'] = val.shortDes
    this.formdata['itemcode'] = val?.FGCode
    this.isShow = true
  }

  saveDataInXl() {
    this.saveSizing('save')
  }

  filterItemCode(val, update) {
    this.filterItemCodeList(val, update, 'ItemCode')
  }

  selectItemCode() {
    if (this.cellItemcode['val']) {
      this.disableitemcodesave = true
      this.saveItemCode('save')
    }
  }

  setColValFromItemCode(meta) {
    if (this.cellItemcode['val']) {
      this.disableitemcodesave = true
      this.setValuesFromItemCode()
    }
  }

  showTechData() {
    this.isShowTechData = true
    this.disableitemcodesave = true
    const y = Number(this.cellItemcode.y)
    const sheet1 = this.$refs.spreadsheet.jexcel
    const instance = sheet1.el
    const productindex = _.findIndex(
      this.jExcelOptions.columns,
      'productcol'
    )
    const currentcell = {
      title: this.jExcelOptions.columns[productindex].title,
      pos: [productindex, y]
    }
    const productname = {
      name: instance.jexcel.getValueFromCoords(productindex, y)
    }

    if (this.cellItemcode.val) {
      setTimeout(() => {
        this.saveItemCode('save')
          .then(() => {
            this.technicalData = true
            this.techQueryVariable = {
              sorting: {
                field: 'id',
                direction: 'ASC'
              },
              paging: { first: 10 },
              filter: {
                makeId: { eq: this.storedummydata.columns.makeId.id },
                ver: { name: { eq: productname.name } },
                companyId: { eq: this.companydata?.id }
              }
            }
          })
      }, 1000)
    } else {
      if (this.companydata.id) { // NOTE old condetion if (this.companydata.id !== 1)
        if (_.isEmpty(productname.name)) {
          this.showNegativeMsg({
            type: 'warning',
            textColor: 'white',
            message: 'Please select Product'
          })
        }

        this.saveTechnicalData(currentcell, productname, false)
          .then(() => {
            this.technicalData = true
            this.disableitemcodesave = false
          })
      }
    }
  }

  onPriceChanged(index, taxRow) {
    this.$set(
      this.offerTax.taxcharges[index],
      'price',
      parseFloat(taxRow?.price || 0)?.toFixed(this.roundCount)
    )

    if (this.offerTax?.format?.id !== 16 && this.taxes) {
      if (this.offerTax?.taxcharges?.length) {
        this.calctotal()
      }
    }
  }

  constructexcelInstance(columns) {
    this.comId = this.companydata.id

    if (this.storedummydata?.columns?.columns) {
      this.storedummydata.columns.columns = _.cloneDeep(columns)
    }

    // NOTE move to user template
    const jexcelwidth = !this.drawerstate.drawer ? '90vw' : '88.5vw'
    const cursystem = this.storedummydata?.measureunit
    const mesurementUnit = this.getmeasureunit
    let origColums = _.cloneDeep(columns)
    const prnuit = mesurementUnit[cursystem]['PR']
    const tempunit = mesurementUnit[cursystem]['Temp']
    const torqunit = mesurementUnit[cursystem]['Torq']
    // _.filter(origColums, {})
    const prArrIndex = _.filter(origColums, z => _.includes(z.title, '_PR')).map(x => x.index)
    const tempArrIndex = _.filter(origColums, z => _.includes(z.title, '_Temp')).map(x => x.index)
    const torqArrIndex = _.filter(origColums, z => _.includes(z.title, '_Torq')).map(x => x.index)

    this.$set(this.jExcelOptions, 'columns', columns)
    this.$set(this.jExcelOptions, 'tableWidth', jexcelwidth)
    this.$set(this.jExcelOptions, 'minDimensions', [columns.length, 50])

    if (
      this.refoptions &&
      this.storedummydata?.proposalRefId &&
      this.interCoComId !== this.company.id
    ) {
      const offerData =
        this.storedummydata?.proposalRefId?.offerSheetData?.offerRowData ||
        this.storedummydata?.proposalRefId?.offerRowData
      this.$worker.postMessage([
        JSON.stringify([offerData, this.refoptions.columns]),
        'create-excel-data-from-json-interCo'
      ])
    }
    const jExcelObj = jspreadsheet(this.$refs.spreadsheet, this.jExcelOptions)
    Object.assign(this, { jExcelObj })
    this.updateproposalRowdata(_.cloneDeep(this.storedummydata))
    const sheet1 = this.$refs.spreadsheet?.jexcel

    if (sheet1) {
      sheet1.ignoreHistory = true
      this.sheetinstace = sheet1
      const stemlpindex = _.findIndex(this.jExcelOptions.columns, 'valveprice')

      if (_.findIndex(this.jExcelOptions.columns, 'productcol') > -1) {
        this.$set(
          sheet1.headers[_.findIndex(this.jExcelOptions.columns, 'productcol')]
            .style,
          'color',
          '#4CAF50'
        )
      }

      if (stemlpindex !== -1) {
        this.$set(sheet1.headers[stemlpindex].style, 'backgroundColor', '#2cc9f5')
      }

      if (!_.isEmpty(this.storedummydata?.offerRowData)) {
        const newsheetMetaData = _.cloneDeep(
          this.storedummydata?.offerRowMetaData
        )
        const offerData = this.storedummydata?.offerRowData
        this.$worker.postMessage([
          JSON.stringify([offerData, this.jExcelOptions.columns, this.company]),
          'create-excel-data-from-json'
        ])
        this.$worker.onmessage = e => {
          if (_.last(e.data) === 'create-excel-data-from-json-interCo') {
            const data = _.first(e.data)
            this.refoptions['data'] = data
            const jExcelObj1 = jexcel(this.$refs.spreadsheet2, this.refoptions)
            Object.assign(this, { jExcelObj1 })
          }

          if (_.last(e.data) === 'create-excel-data-from-json') {
            const data = _.first(e.data)
            this.$set(sheet1.options, 'meta', newsheetMetaData)
            sheet1.setData(data, this)

            if (offerData.length <= 300) {
              this.setAdditionalXLData(offerData, sheet1, JSON.stringify(columns), this.company)
            }
          }
      }
      } else this.toggleSpinner(false)
    } else this.toggleSpinner(false)

    this.orignalcolumns = _.cloneDeep(this.jExcelOptions.columns)

  }

  columnindex() {
    const alphabet = []
    for (let j = 0; j < this.jExcelOptions.columns.length; j++) {
      const m = 65 + j
      let k = m % 65
      k = k > 25 ? k % 26 : k
      if (m < 91) {
        alphabet.push(String.fromCharCode(k + 65))
      } else if (m > 90 && m < 117) {
        alphabet.push(String.fromCharCode(65) + String.fromCharCode(k + 65))
      } else if (m > 117 && m < 143) {
        alphabet.push(String.fromCharCode(66) + String.fromCharCode(k + 65))
      }
    }
    this.alphabetlist = alphabet
  }

  mounttaxdata(val) {
    return new Promise(resolve => {
      this.taxpriceval = val
      const sheetinstace = this.sheetinstace
      const qtycol = _.find(this.jExcelOptions.columns, 'qty')
      const lumpsumval = _.find(sheetinstace?.options?.columns, 'testing')
      // console.log(this.storedummydata, 'this.storedummydata', this.companydata)
      this.$worker.postMessage([
        JSON.stringify([
          val,
          this.companydata,
          this.storedummydata,
          qtycol,
          lumpsumval,
          this.discountprice
        ]),
        'mount-tax-data'
      ])
      this.$worker.onmessage = e => {
        if (_.last(e.data) === 'mount-tax-data') {
          resolve(_.first(e.data))
        }
      }
    })
  }

  calctotal(returnval = false, taxCharges) {
    if(taxCharges || this.offerTax?.taxcharges) {
    const self = this
    const companyInvoice = _.cloneDeep(this.companydata?.invoiceCharge)
    const taxList = returnval ? taxCharges : self.offerTax?.taxcharges
    const grandTotalIndex = _.findIndex(
      returnval ? taxCharges : this.offerTax?.taxcharges,
      'grandTotal'
    )
    const grandTotalvalue = parseFloat(_.sum(
      _.map(taxList, v => {
        if (!v.grandTotal) {
          if (companyInvoice?.discountColumn) {
            if (!v.isamount) {
              return parseFloat(v.price)
            }
          } else {
            return parseFloat(v.price)
          }
        }
      }).filter(val => val)
    ) || 0)?.toFixed(this.roundCount)

    if (returnval) {
      return grandTotalvalue
    } else {
      if (this.offerTax.taxcharges[grandTotalIndex]) {
        this.$set(
          this.offerTax.taxcharges[grandTotalIndex],
          'price',
          grandTotalvalue
        )
      }
    }
  }
  }

  changetaxtype(taxtype) {
    taxtype = taxtype?.toLowerCase()
    const companyInvoice = _.cloneDeep(this.companydata?.invoiceCharge)
    let taxFields = _.filter(this.offerTax?.taxcharges, z => !z.taxfield)
    const newFields = _.map(companyInvoice?.taxformatObj?.[taxtype], x => {
      if (x) {
        x.price = parseFloat(
          this.calculatedvalue(x, { returnval: true })
        )?.toFixed(this.roundCount)
      }

      return x
    }).filter(val => val)

    this.$nextTick(() => {
      // console.log(newFields, companyInvoice?.taxformatObj, taxFields)
      taxFields.splice(taxFields.length - 1, 0, ...newFields)
      const otherCharges = _.findIndex(
        taxFields,
        ({ name }) => _.includes(_.toLower(name), 'other charges')
      )
      const igstIndex = _.findIndex(
        taxFields,
        ({ name }) => _.includes(_.toLower(name), 'gst')
      )
      const sgstIndex = _.findIndex(
        taxFields,
        ({ name }) => _.includes(_.toLower(name), 'sgst')
      )
      const totalAmountPrice = _.find(
        taxFields,
        ({ name }) => _.includes(_.toLower(name), 'total amount')
      )?.price || 0
      const totalAmountIndex = _.findIndex(
        this.offerTax.taxcharges,
        ({ name }) => _.includes(_.toLower(name), 'total amount')
      )

      if (otherCharges > -1 && taxtype === 'interstate') {
        taxFields = this.arrayMove(taxFields, igstIndex, otherCharges)
      } else if (otherCharges > -1) {
        taxFields = this.arrayMove(taxFields, otherCharges, sgstIndex)
      }
      this.$nextTick(() => {
        this.$set(this.offerTax, 'taxcharges', taxFields.filter(val => val))

        _.forEach(this.offerTax.taxcharges, (taxRow, index) => {
          if (taxRow) {
            this.calculatedvalue(taxRow, { template: true, index })
          }
        })
        // this.calctotal()
        this.$nextTick(() => {
          this.offerTax.taxcharges[totalAmountIndex].price = Number(totalAmountPrice)?.toFixed(this.roundCount)
        })
      })
    })
  }

  calculatedvalue(val, options) {
    let priceval = 0
    let frieght = 0
    let packaging = 0

    if (val.price) {
      val.price = Number(val.price)
    }
    const gstIndex = _.findIndex(
      this.offerTax.taxcharges,
      ({ name }) => _.includes(_.toLower(name), 'gst')
    )
    const sgstIndex = _.findIndex(
      this.offerTax.taxcharges,
      ({ name }) => _.includes(_.toLower(name), 'sgst')
    )
    const companyInvoice = _.cloneDeep(this.companydata?.invoiceCharge)
    this.$nextTick(() => {
      if (val) {
        if (_.includes(_.toLower(val.name), 'freight')) {
          frieght = _.round(val.price || 0, this.roundCount)
        }
        if (_.includes(_.toLower(val.name), 'packaging')) {
          packaging = _.round(val.price || 0, this.roundCount)
        }
        if (
          _.includes(_.toLower(val.name), 'gst') ||
          _.includes(_.toLower(val.name), 'sgst') ||
          _.includes(_.toLower(val.name), 'freight') ||
          _.includes(_.toLower(val.name), 'packaging')
        ) {
          if (!frieght) {
            frieght = _.find(
              this.offerTax.taxcharges,
              (row) => _.includes(_.toLower(row.name), 'freight')
            )?.price || 0
          }
          if (!packaging) {
            packaging = _.find(
              this.offerTax.taxcharges,
              (row) => _.includes(_.toLower(row.name), 'packaging')
            )?.price || 0
          }
        }
      }
    })

    if (val.isdiscountcol) {
      const qtycol = _.find(this.jExcelOptions.columns, 'qty')
      const sheetdata = _.cloneDeep(this.storedummydata?.offerRowData)
      const lumpsumval = _.find(this.sheetinstace?.options?.columns, 'testing')
      priceval = _.round(
        Number(this.taxpriceval.discountamount) -
          Number(val.percent) * (Number(this.taxpriceval.discountamount) / 100) +
          (sheetdata.length && lumpsumval?.title && sheetdata?.[0]?.[qtycol.title]
            ? _.sum(
              sheetdata.map(
                x => Number(x[lumpsumval?.title]) * Number(x[qtycol.title] || 0)
              )
            )
            : 0),
        0
      )
      this.discountprice = priceval
    } else {
      priceval = companyInvoice?.discountColumn
        ? Number(val?.percent || 0) * (Number(this.discountprice) / 100)
        : Number(val?.percent || 0) * (Number(this.taxpriceval?.totalamount) / 100)
        this.$nextTick(() => {
          const sumOthers =  _.sum(_.filter(this.offerTax.taxcharges, ({ name }) => (!_.includes(['igst', 'gst', 'cgst', 'sgst','total amount', 'grand total value', 'sgst/ugst', 'amount after discount'], _.toLower(name)))).map(x=> Number(x.price)))
          const gstVal = companyInvoice?.discountColumn
            ? Number(this.offerTax.taxcharges[gstIndex]?.percent) * ((Number(this.discountprice) + Number(sumOthers)) / 100)
            : Number(this.offerTax.taxcharges[gstIndex]?.percent) * ((Number(this.taxpriceval?.totalamount) + Number(sumOthers)) / 100)
          const sgstVal = companyInvoice?.discountColumn
            ? Number(this.offerTax.taxcharges[sgstIndex]?.percent) * ((Number(this.discountprice) + Number(sumOthers)) / 100)
            : Number(this.offerTax.taxcharges[sgstIndex]?.percent) * ((Number(this.taxpriceval?.totalamount) + Number(sumOthers)) / 100)

          if (_.includes(_.toLower(val.name), ['gst', 'igst'])) {
            priceval = _.round(gstVal, this.roundCount)
          }

          if (this.offerTax.taxcharges?.[gstIndex]) {
            this.$set(
              this.offerTax.taxcharges[gstIndex],
              'price',
              parseFloat(gstVal || 0).toFixed(this.roundCount)
            )
          }
          if (this.offerTax.taxcharges?.[sgstIndex]) {
            this.$set(
              this.offerTax.taxcharges[sgstIndex],
              'price',
              parseFloat(sgstVal || 0).toFixed(this.roundCount)
            )
          }
          this.calctotal()
          // this.offerTax.taxcharges[totalAmountIndex].price = Number(totalAmountPrice)?.toFixed(this.roundCount)
        })
      }

    if (options.returnval) {
      return priceval
    } else {
      if (val.isdiscountcol) {
        this.offerTax.taxcharges = _.map(this.offerTax?.taxcharges, a => {
          if (a) {
            if (!a.default) {
              a.price = parseFloat(
                this.calculatedvalue(a, { returnval: true })
              )?.toFixed(this.roundCount)
            } else if (a.isdiscountcol) {
              a.price = parseFloat(
                this.discountprice || 0
              )?.toFixed(this.roundCount)
            }
          }

          return a
        }).filter(val => val)

        this.$nextTick(() => {
          this.calctotal()
        })
      } else if (options.template) {
        if (priceval) {
        this.$set(
          this.offerTax.taxcharges[options?.index],
          'price',
          _.round(priceval, 0)
        )
        }
        this.$nextTick(() => {
          this.calctotal(false, this.offerTax.taxcharges)
        })
      }
    }
  }

  /**
   * @method getDesscription
   * @return {Promise}
   */
  getDesscriptionByItemCode(itemCode) {
    const QUERY = gql`
      query getDescriptionByItemCode(
        $orgId: Int!
        $itemCode: String!
      ) {
        descriptionByItemCode: getDescriptionByItemCode(
          orgId: $orgId
          itemCode: $itemCode
        )
      }
    `
    return this.$apollo.query({
      query: QUERY,
      variables: {
        itemCode,
        orgId: this.getProposaldata.orgId
      }
    })
  }

  /**
   * @method onItemCodeInput
   * @param {string} itemCode
   * @return {void}
   */
  onItemCodeInput(itemCode) {
    this.formdata['itemcode'] = itemCode ? itemCode.id : null
    this.setPrice(this.formdata['itemcode'], true)
    if (!this.curColumn?.productId?.groupId.name.toLowerCase()?.includes('accessories')) {
      if (this.formdata['itemcode']) {
        this.getDesscriptionByItemCode(this.formdata['itemcode'])
          .then(({ data: { descriptionByItemCode } }) => {
            this.$set(this.formdata, 'desc', descriptionByItemCode?.description || '')

            if (!descriptionByItemCode) {
              this.showInfoMsg({
                textColor: 'white',
                message: 'No Desctiption found for the selected itemcode.!'
              })
            }
          })
      } else {
        this.formdata['desc'] = ''
      }
    }
  }

  savetax() {
    this.taxes = false
    this.toggleSpinner(true)
    const propdata = this.storedummydata
    propdata['taxData'] = this.offerTax
    this.updateproposalRowdata(propdata)
    this.saveOfferSheet(
      _.cloneDeep({ ...this.currentsheetData, taxData: this.offerTax }),
      false,
      [],
      this.jExcelOptions.column
    ).then(({ data }) => {
      if (!data?.offerSheet?.id) {
        this.showWarningMsg({
          textColor: 'white',
          message: 'Check your network connection. Please try again.!'
        })
        this.toggleSpinner(false)
        return
      }

      if (!_.isEmpty(data?.offerSheet)) {
        this.$nextTick(() => {
          this.setOfferId(Number(data?.offerSheet?.id))
          this.updateproposalRowdata({
            ...this.storedummydata,
            offerSheetId: data.offerSheet.id
          })
          this.$emit('nextpage', true)
          this.showPositiveMsg({ message: 'Record updated successfully!' })
        })
      }
      this.toggleSpinner(false)
    }).catch((error) => {
      this.toggleSpinner(false)
      // console.error(error)
    })
  }

  importsheetdata(newsheet, sheetcomment, otherSheet = {}) {
    const proposal = _.merge(_.cloneDeep(this.storedummydata), otherSheet)
    this.updateproposalRowdata(proposal)
      .then(() => {
        this.storedummydata = _.cloneDeep(proposal)

        if (window.Worker) {
          const sheet = this.$refs.spreadsheet.jexcel

          // NOTE deleteing existing sheet data
          sheet.deleteRow(0, sheet.options.data.length - 1)

          // NOTE set initial data worker
          this.$worker.postMessage([
            newsheet,
            sheetcomment,
            JSON.stringify(this.getProposaldata.columns),
            JSON.stringify(sheet.options.columns),
            this.company,
            true,
            'set-initial-data'
          ])
          this.$worker.onmessage = e => {
            if (_.last(e.data) === 'set-initial-data') {
              if (_.isArray(e.data)) {
                const [from, method, data] = e.data

                if (from === 'codeCreation') {
                  const [title, index, val] = data

                  switch (method) {
                    case 'setMeta':
                      if (!_.isEmpty(val)) {
                        sheet[method](title, index, val)
                      }
                      break
                  }
                }

                if (from === 'finalData') {
                  this.jExcelObj.options.meta['fgCodeType'] = data.fgCodeType
                  sheet[method](data.sheetVal, this)

                  if (newsheet.length <= 300) {
                    this.setAdditionalXLData(
                      sheet.getJson(),
                      sheet,
                      JSON.stringify(sheet.options.columns),
                      this.company,
                      true
                    )
                  } else {
                    this.$emit('sumdata', data.finalData)
                  }
                }
              }
            }
          }
        } else {
          this.toggleSpinner(false)
        }
      })
  }

  setAdditionalXLData(newsheet, sheet, columns, company, isFromImport = false) {
    const sheetmeta = JSON.stringify(sheet.getMeta() || {})
    this.$worker.postMessage([
      newsheet,
      columns,
      sheetmeta,
      company,
      isFromImport,
      'set-additional-xl-data'
    ])
    this.$worker.onmessage = e => {
      if (_.last(e.data) === 'set-additional-xl-data') {
        if (_.isArray(e.data)) {
          const [operation, action, data] = e.data

          if (operation === 'set-data') {
            if (!this.notif) {
              this.$emit('setting-additional-xl-data', true)
              this.notif = this.$q.notify({
                timeout: 0, // we want to be in control when it gets dismissed
                group: false, // required to be updatable
                spinner: true,
                caption: '0%',
                position: 'top-right',
                message: 'Setting up XL addetional data...'
              })
            }
            this.notifPercentage += 1
            this.notif({ caption: `${this.notifPercentage}` })

            let targetCell = ''
            const [indexArray, value, isLastIndex] = data

            if (_.isArray(indexArray)) {
              const columnNameTarget = jexcel?.getColumnNameFromId(indexArray)
              targetCell = sheet?.getCell(columnNameTarget)
            }

            if (isLastIndex) {
              this.notif({
                icon: 'done', // we add an icon
                spinner: false, // we reset the spinner setting so the icon can be displayed
                message: 'Uploading almost done!',
                timeout: 2500 // we will timeout it in 2.5s
              })
              setTimeout(() => {
                this.notif = null
                this.notifPercentage = null
                this.$emit('setting-additional-xl-data', false)
              }, 1000)
            }

            switch (action) {
              case 'setInnerText':
                if (
                  targetCell &&
                  value !== undefined &&
                  value !== 'undefined'
                ) {
                  targetCell.innerText = value
                }
                break
              case 'setComment':
                if (value?.length) {
                  sheet.setComments(indexArray, value)
                }
                break
              case 'readonly':
                if (targetCell) {
                  targetCell.classList.add('readonly')
                }
                break
            }

            if (isLastIndex) {
              this.$emit('setting-additional-xl-data', false)
            }
          }

          if (operation === 'finalData') {
            this.$emit('sumdata', data)
          }
        }
      }
    }
  }

  gettechdata(products, productlist, productindex) {
    const storedata = _.cloneDeep(this.getProposaldata)
    const queryfilter = {
      ver: { name: { in: products } },
      companyId: { eq: this.companydata.id },
      makeId: { eq: storedata.columns.makeId.id }
    }
    this.$apollo.query({
      query: masterTechData,
      variables: { filter: queryfilter }
    }).then(({ data }) => {
      if (!_.isEmpty(data)) {
        const node = _.map(data?.masterTeches?.edges, 'node')
        const techdata = _.cloneDeep(node)
        const column = this.jExcelOptions.columns[productindex]

        if (!storedata.technicaldata) {
          storedata['technicaldata'] = {}
        }

        _.forEach(productlist, (p, i) => {
          const technicaldata = _.find(techdata, ['verId.name', p])

          if (technicaldata) {
            storedata.technicaldata[`row_${i}`] = [
              {
                masterdata: technicaldata,
                verId: technicaldata?.verId,
                dataSrc: technicaldata.dataSrc,
                column: column.title,
                portfolioId: technicaldata.portfolioId
              }
            ]
          }
        })

        if (!storedata.technicaldata) {
          storedata['technicaldata'] = {}
        }

        _.forEach(productlist, (p, i) => {
          const technicaldata = _.find(techdata, ['verId.name', p])

          if (technicaldata) {
            storedata.technicaldata[`row_${i}`] = [
              {
                masterdata: technicaldata,
                verId: technicaldata?.verId,
                dataSrc: technicaldata.dataSrc,
                column: column.title,
                portfolioId: technicaldata.portfolioId
              }
            ]
          }
        })

        this.updateproposalRowdata(storedata)
      }
    })
  }

  async actuatorTorqueValue() {
    const val = this.actsizingdata

    if (!val?.valveType || !val?.diffPr || !val?.makeId) {
      this.showWarningMsg({
        message: 'Select ValveType, diffrential Pressure, Make.!',
        textColor: 'white'
      })
      return
    }

    const series = this.currentcell.instance.jexcel.getValueFromCoords(
      _.findIndex(this.jExcelOptions.columns, ['sourceId.name', 'SERIES']),
      this.currentcell.pos[1]
    )

    const makeId = Number(val.makeId) // === 3 ? 2 : Number(val.makeId)
    const variables = { makeId }

    if (makeId === 2 && this.company.code === '08_DELVAL_USA') {
      variables['makeId'] = 3
    } else if (makeId === 3 && this.company.code === '03_DELVAL') {
      variables['makeId'] = 2
    }

    if (!val?.valveType?.id) {
      variables['productName'] = val.valveType?.name
    } else {
      variables['productId'] = val.valveType.id
    }

    if (_.includes(series, ':-')) {
      variables['series'] = series.split(':-')[1]
    }

    this.$apollo.query({
      variables,
      // fetchPolicy: 'network-only',
      query: queries.getseriestorqueList
    }).then(({ data: { getTorqueListForSeries } }) => {
      if (getTorqueListForSeries?.length) {
        let y
        const cell = this.currentcell
        const x = this.actsizingdata.diffPr
        const [zero, one] = _.split(series, ':-')
        const makeId = this.actsizingdata['makeId']
        this.getTorqueListForSeries = getTorqueListForSeries
        const list = _.find(getTorqueListForSeries, ['series.code', one !== '' ? one : zero])
        const splittedArray = _.map(list.differentialPressure, (row) => _.split(row, ':-'))
        const differentialPressure = _.map(
          splittedArray,
          ([zero, one]) => ({ value: Number(makeId === 3 ? zero : one), code: one })
        )
        const valExist = _.find(differentialPressure, ({ value }) => value === x)
        const lowerVal = _.maxBy(
          _.filter(differentialPressure, ({ value }) => value < x),
          'value'
        )
        const higherVal = _.minBy(
          _.filter(differentialPressure, ({ value }) => value > x),
          'value'
        )

        const getCode = (diffPr) => {
          const tkey = [one !== '' ? one : zero]
          _.forEach(list?.torqueColumn, (col) => {
            if (Number(col?.sourceId?.id || 0) === 148) {
              tkey.push(diffPr)
            } else if (_.isArray(col.columns)) {
              const condetion = {}

              _.forEach(col.columns, (derivedCol) => {
                if (derivedCol?.sourceId) { // NOTE if it is derived column
                  let columnIndex = _.findIndex(this.jExcelOptions.columns, [
                    'sourceId.id',
                    derivedCol.sourceId.id
                  ])

                  if (columnIndex < 0) {
                    columnIndex = _.findIndex(this.jExcelOptions.columns, [
                      'sourceId.id',
                      Number(derivedCol.sourceId.id)
                    ])
                  }

                  const cellval = cell.instance.jexcel.getValueFromCoords(
                    columnIndex,
                    cell.pos[1]
                  )
                  const [value0] = _.split(cellval, ':-')

                  condetion[_.toLower(derivedCol.sourceId?.name)] = value0
                }
              })

              const fgtCode = _.find(col.source, condetion)?.fgtCode || 0

              tkey.push(fgtCode)
            } else if (col?.sourceId) {
              const cellval = cell.instance.jexcel.getValueFromCoords(
                _.findIndex(this.jExcelOptions.columns, [
                  'sourceId.id',
                  Number(col.sourceId.id)
                ]),
                cell.pos[1]
              )

              const value = cellval.split(':-')
              tkey.push(value[1] === '' ? value[0] : value[1])
            }
          })

          return tkey
        }

        if (_.isEmpty(valExist)) {
          const x1 = lowerVal ? Number(lowerVal.value) : null
          const x2 = higherVal ? Number(higherVal.value) : null
          const x1CodeArr = getCode(lowerVal?.code)
          const x2CodeArr = getCode(higherVal?.code)
          const x1Code = x1CodeArr.join('')
          const x2Code = x2CodeArr.join('')
          const y1 = Number(list?.torqueList?.[x1Code]?.BTO || 0)
          const y2 = Number(list?.torqueList?.[x2Code]?.BTO || 0)

          if (!y2) {
            y = y1
          }

          if (!y1 && y2) {
            y = y2
          }

          if (y1 && y2) {
            y = _.round(y1 + (x - x1) * (y2 - y1) / (x2 - x1), 0)
          }
        } else {
          const codeArr = getCode(valExist.code)
          y = Number(list?.torqueList?.[codeArr.join('')]?.BTO || 0)
        }

        if (y) {
          if (makeId === 2 && this.company.code === '08_DELVAL_USA') {
            y *= 0.113
          }

          this.$set(this.actsizingdata, 'vtorque', Number(y) || 0)
        } else {
          this.showWarningMsg({
            message: `Please check following fields ${_.join(
              _.map(_.first(getTorqueListForSeries)?.torqueColumn, 'title')
            )}`,
            textColor: 'white'
          })
          this.$set(this.actsizingdata, 'vtorque', 0)
        }
      }
    })
  }

  showitemcode(series) {
    if (series) {
      itemcodelist = []
      this.code = true
      const priceList = _.cloneDeep(this.getPriceList)
      this.currentPriceList = _.find(priceList, {
        companyId: this.companydata?.id,
        series: { name: series?.name }
      })
      if (
        this.currentPriceList &&
        _.includes(this.currentPriceList.orgId, this.storedummydata?.orgId)
      ) {
        itemcodelist = this.currentPriceList.priceList
        const y = this.currentcell.pos[1] // col, row

        const sheet1 = this.$refs.spreadsheet.jexcel
        const prodTypeIndex = _.findIndex(this.jExcelOptions.columns, { id: 'Product_Type'})
        if ( prodTypeIndex) {
          const cellvalue = sheet1.getValueFromCoords(
            prodTypeIndex,
            y
          )
          let strfilter = ''
          if (_.includes(['Automated', 'Bare'], cellvalue)) {
            strfilter = 'G-0'
            itemcodelist =  _.pickBy(itemcodelist, (value, key) => {
              return !_.endsWith(key, strfilter);
            })
          } else if (cellvalue === 'Manual') {
            strfilter = 'A-0'
            itemcodelist =  _.pickBy(itemcodelist, (value, key) => {
              return !_.endsWith(key, strfilter);
            })
          }

      }
      } else {
        itemcodelist = []
        this.getitemcodes(this.storedummydata.columns.makeId, {
          field: 'series',
          value: series
        })
      }
    }
  }

  getitemcodes(make, optfield, acessories = false) {
    let optvar = {}
    let actSeries =  null
    if(this.actSizing) {
      actSeries = this.actsizingdata?.series
    }
    const filterval = {
      productId: null,
      makeId: Number(make?.id),
      companyId: Number(this.companydata.id),
      orgId: !this.storedummydata.orgId
        ? Number(this.salesRegion.id)
        : this.storedummydata.orgId,
      seriesId: actSeries
    }

    if (optfield.field === 'series') {
      optvar = filterval

      if (optfield.value?.id) {
        optvar['seriesId'] = optfield.value?.id
      }
    } else {
      optvar = filterval // { ...filterval, [optfield.field]: Number(optfield.value.id) }

      if (optfield?.value?.id) {
        optvar[optfield.field] = Number(optfield.value.id)
      }

      if (optvar.productId === null) return
    }

    const query = acessories
      ? queries.getAccessoriespricelist
      : queries.getproposalpricelist

    if (optvar?.seriesId) {
      const [name, code] = _.split(optvar?.seriesId, ':-')

      optvar.seriesId = `:-${code}`
    }

    this.$apollo.query({
      query,
      variables: optvar
      // fetchPolicy: 'network-only'
    }).then(({ data }) => {
      const edges = data.getPriceList

      if (edges?.length) {
        const list = Object.values(edges.flatMap(z => z.priceList)).flat()
        itemcodelist = {}
        _.forEach(list, x => {
          itemcodelist = { ...itemcodelist, ...x }
        })

        const y = this.currentcell.pos[1] // col, row
        const sheet1 = this.$refs.spreadsheet.jexcel
        const prodTypeIndex = _.findIndex(this.jExcelOptions.columns, { id: 'Product_Type'})
        if ( prodTypeIndex) {
          const cellvalue = sheet1.getValueFromCoords(
            prodTypeIndex,
            y
          )
          let strfilter = ''
          if (_.includes(['Automated', 'Bare'], cellvalue)) {
            strfilter = 'G-0'
            itemcodelist =  _.pickBy(itemcodelist, (value, key) => {
              return !_.endsWith(key, strfilter);
            })
          } else if (cellvalue === 'Manual') {
            strfilter = 'A-0'
            itemcodelist =  _.pickBy(itemcodelist, (value, key) => {
              return !_.endsWith(key, strfilter);
            })
          }
        }

        this.itemcodeexceldata = acessories
          ? _.flatMap(edges, 'excelSheetData')
          : []
        const firstData = _.first(edges)
        const obj = {
          productId: firstData?.productId,
          companyId: firstData?.companyId,
          id: firstData?.id,
          makeId: firstData?.makeId,
          orgId: firstData?.orgId,
          priceList: itemcodelist,
          excelSheetData: this.itemcodeexceldata
        }

        if (firstData?.series) {
          obj['series'] = firstData?.series
        }
        // NOTE update priceList store
        this.addPriceList(obj)
      }
    })
  }

  /**
   * @method setnsDependency
   * @param { String } val
   * @param { Method } update
   * @returns { Void }
   */
  setnsDependency() {
    if (this.nsText?.dependentvalues) {
      this.nsDependentvalue = this.nsText?.dependentvalues || null
    }
  }

  createNewDiffpr(val, done) {
    if (val?.length) {
      const listValue = _.find(this.ListOfDiffPressure, ['value', val])
      if (!listValue) {
        const masterFieldValue = {
          value: val,
          masterFieldId: 148,
          companyId: [this.company.id]
        } // NOTE 148 => differential Pressure
        this.$apollo
          .mutate({
            mutation: mutations.createMasterFieldValue,
            variables: { input: { masterFieldValue } }
          })
          .then(({ data: { masterValue } }) => {
            this.ListOfDiffPressure.push(masterValue)
            done(masterValue, 'toggle')
          })
      } else {
        done(listValue, 'toggle')
      }
    }
  }

  filterseries(val, update) {
    const series = _.cloneDeep(
      _.find(this.jExcelOptions.columns, {
        sourceId: { name: 'SERIES' }
      })
    )
    const sheet1 = this.$refs.spreadsheet.jexcel

    if (series?.productDependent) {
      const cellvalue = sheet1.getValueFromCoords(
        _.findIndex(this.jExcelOptions.columns, 'productcol'),
        this.cellItemcode.y
      )
      if (cellvalue.length) {
        series.source = _.filter(series.source, x =>
          _.includes(x.Product, cellvalue)
        )
      }
    }

    update(() => {
      if (val === '') {
        this.serieslist = series?.source || []
      } else {
        const needle = _.toLower(val)
        this.serieslist = series.source
          .filter(v => v.name.toLowerCase().indexOf(needle) > -1)
          .slice(0, 9)
      }
    })
  }

  filterItemCodeList(val, update, columntype, removeDash = true) {
    let returnItemcode = Object.keys(itemcodelist)

    if (removeDash) {
      returnItemcode = _.map(
        Object.keys(itemcodelist),
        (row) => ({ code: _.replace((row || ''), /-/g, ''), id: row })
      )
    }

    if (columntype === 'ItemCode') {
      update(() => {
        if (val === '') {
          this.listOfItemcode = returnItemcode
        } else {
          const needle = val.toLowerCase()
          this.listOfItemcode = returnItemcode.filter(
            v => (
              removeDash
                ? v?.code?.toLowerCase()?.indexOf(needle) > -1
                : v?.toLowerCase()?.indexOf(needle) > -1
            )
          )
        }
      })
    } else {
      filteredlist = this.formdata.itemcode?.length
        ? _.map(
          _.filter(this.itemcodeexceldata, ['itemcode', this.formdata.itemcode]),
          'itemcode'
        )
        : []

      if (removeDash) {
        filteredlist = _.map(
          filteredlist,
          (row) => ({ code: _.replace((row || ''), /-/g, ''), id: row })
        )
      }

      update(() => {
        if (val === '') {
          this.listOfItemcode = !filteredlist?.length
            ? returnItemcode
            : filteredlist
        } else {
          const needle = val.toLowerCase()
          if (_.isEmpty(filteredlist)) {
            this.listOfItemcode = _.filter(
              returnItemcode,
              v => (
                removeDash
                  ? v?.code?.toLowerCase()?.indexOf(needle) > -1
                  : v?.toLowerCase()?.indexOf(needle) > -1
              )
            )
          } else {
            this.listOfItemcode = _.filter(
              filteredlist,
              v => (
                removeDash
                  ? v?.code?.toLowerCase()?.indexOf(needle) > -1
                  : v?.toLowerCase()?.indexOf(needle) > -1
              )
            )
          }
        }
      })
    }
  }

  /**
   * @method Company
   * @param { String } val
   * @param { Method } update
   * @returns { Void }
   */
  filterProductName(val, update) {
    const variables = {
      filter: { name: { iLike: `%${val}%` }, groupId: { in: [1, 4] } }
    }
    this.filterFunction(val, update, {
      query: queries.getProductVersions,
      variables,
      field: 'ListOfProductName',
      resultName: 'productVersions'
    })
  }

  /**
   * @method Company
   * @param { String } val
   * @param { Method } update
   * @returns { Void }
   */
  filterActuator(val, update) {
    const variables = {
      filter: {
        name: { iLike: `%${val}%` },
        groupId: { in: [5, 6] },
        comId: { eq: this.company.id }
      }
    }
    this.filterFunction(val, update, {
      query: queries.getProductVersions,
      variables,
      field: 'ListOfActuator',
      resultName: 'productVersions'
    })
  }

  /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
  filterFailSafe(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 150,
      field: 'ListOfFailSafeCondition',
      resultName: 'masterFieldValues'
    })
  }


  /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
   filterSupplyVoltage(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 72,
      field: 'ListOfSupplyVoltage',
      resultName: 'masterFieldValues'
    })
  }



  /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
   filterHertz(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 135,
      field: 'ListOfHertz',
      resultName: 'masterFieldValues'
    })
  }

  /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
   filterPhases(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 133,
      field: 'ListOfPhase',
      resultName: 'masterFieldValues'
    })
  }

    /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
   filterSpecial(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 124,
      field: 'ListOfSpecial',
      resultName: 'masterFieldValues'
    })
  }

    /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
   filterInsert(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 50,
      field: 'ListOfInsert',
      resultName: 'masterFieldValues'
    })
  }

    /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
   filterEnclosure(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 136,
      field: 'ListOfEnclosure',
      resultName: 'masterFieldValues'
    })
  }

      /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
   filterMOR(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 82,
      field: 'ListOfMOR',
      resultName: 'masterFieldValues'
    })
  }

  /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
  filterMediaType(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 149,
      field: 'ListOfMediaType',
      resultName: 'masterFieldValues'
    })
  }

  /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
  filterDiffPressure(val, update) {
    return this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 148,
      field: 'ListOfDiffPressure',
      resultName: 'masterFieldValues'
    })
  }

  /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
  filterSafetyFactor(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 152,
      field: 'ListOfSafetyFactor',
      resultName: 'masterFieldValues'
    })
  }

  /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
  filterSupplyPressure(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 151,
      field: 'ListOfSupplyPressure',
      resultName: 'masterFieldValues'
    })
  }

  /**
   * @method MasterFieldValues
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
  filterActuatorType(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 153,
      field: 'ListOfActuatorType',
      resultName: 'masterFieldValues'
    })
  }

  /**
   * @method filterActuatorMounting
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
  filterActuatorMounting(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 154,
      field: 'ListOfActuatorMounting',
      resultName: 'masterFieldValues'
    })
  }

  /**
   * @method filterManualOverride
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
  filterManualOverride(val, update) {
    this.commonFilterUpdater(val, update, {
      masterField: true,
      noLengthCheck: true,
      masterFieldId: 155,
      field: 'ListOfManualOverride',
      resultName: 'masterFieldValues'
    })
  }

  /**
   * @method filtermodel
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
  filtermodel(val, update) {
    let returndata = _.filter(_.map(this.itemcodeexceldata, 'Model'), val => val)

    if (_.isEmpty(returndata)) {
      returndata = [this.formdata?.model]
    }

    returndata = _.uniq(returndata)

    update(() => {
      if (val === '') {
        this.modellist = returndata
      } else {
        const needle = val.toLowerCase()
        this.modellist = _.filter(
          returndata,
          v => v.toLowerCase().indexOf(needle) > -1
        )
      }
    })
  }

  /**
   * @method filternsList
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
  filternsList(val, update) {
    const needle = val.toLowerCase()
    const filter = {
      name: { iLike: `%${val}%` },
      companyId: { eq: this.companydata.id },
      makeId: { eq: Number(this.storedummydata.columns.makeId?.id) },
      sourceId: { eq: Number(this.nsColumnNameTarget.sourceId?.id) }
    }

    if (!this.nsColumnNameTarget.sourceId?.id) {
      this.showInfoMsg({ message: 'Source ID Not found for the selected column. No NS data will be loaded.!' })
    }

    if (
      this.nsColumnNameTarget.sourceId?.id &&
      this.storedummydata.columns.makeId?.id
    ) {
      this.$apollo.query({
        query: queries.getNonStandards,
        variables: {
          filter,
          sorting: {
            field: 'id',
            direction: 'ASC'
          },
          paging: { first: 20 }
        },
        fetchPolicy: 'network-only'
      }).then(({ data }) => {
        const edges = data?.nonStandards?.edges
        nsOptions = _.map(edges, 'node')

        update(() => {
          if (val === '') {
            this.nonStandardList = _.cloneDeep(nsOptions)
          } else {
            this.nonStandardList = _.filter(
              nsOptions,
              v => v?.name?.toLowerCase()?.indexOf(needle) > -1
            )
          }
        })
      })
    }
  }

  /**
   * @method filternsdependent
   * @param { String } val
   * @param { Function } update
   * @return { Void }
   */
  filternsdependent(val, update) {
    const needle = _.toLower(val)
    update(() => {
      if (val === '') this.nsDependendlist = nsDependentopt
      else {
        this.nsDependendlist = _.filter(
          nsDependentopt,
          x => x?.name?.toLowerCase()?.indexOf(needle) > -1
          // _.indexOf(_.toLower(x?.name), needle) > -1
        )
      }
    })
  }

  /**
   * @method commonFilterUpdater
   * @param { String } val
   * @param { Function } update
   * @param { Object } options
   * @return { Void }
   */
  commonFilterUpdater(val, update, options) {
    return new Promise((resolve) => {
      let query
      const variables = {
        filter: {},
        sorting: {
          field: '',
          direction: 'ASC'
        },
        paging: { first: 500 }
      }

      if (options.masterField) {
        query = queries.getMasterFieldValueForDropDown
        variables.sorting.field = 'value'
        variables.filter['value'] = { iLike: `%${val}%` }
        variables.filter['masterFieldId'] = { eq: options.masterFieldId } // ? master_field_id => 1
      }

      if (
        (val && val.length > 0) ||
        options.noLengthCheck ||
        (val && val.length > options.minLength)
      ) {
        this.gqlQuery(query, variables, options.resultName)
          .then(result => {
            console.log(this.actsizingdata, 'this.actsizingdata')
            let verId =  this.actsizingdata['actuator']?.id === 185 ? 10 : this.actsizingdata['actuator']?.id
            if (verId === 193) {
              verId = 9
            }
            if (result[0]?.data.makeId && _.includes([151], options.masterFieldId)) {
              result = _.filter(result, [
                'data.makeId.id',
                this.actsizingdata['makeId']
              ])
            }
            if (result[0]?.data.makeId && _.includes([124, 82, 50], options.masterFieldId)) {
              result = _.filter(result, [
                'data.verId',
                verId
              ])
            }
            if (update) {
              update(() => {
                this[options.field] = _.uniqBy(result, 'value')
              })
            } else {
              this[options.field] = _.uniqBy(result, 'value')
            }
            resolve()
          })
      } else {
        this[options.field] = []
        update()
        resolve()
      }
    })
  }

  createNSvalue(val, done, dependency) {
    if (val.length > 0) {
      if (!_.map(nsOptions, 'name').includes(val)) {
        const obj = {
          name: val,
          sourceId: this.nsColumnNameTarget?.sourceId,
          companyId: this.companydata.id,
          makeId: this.storedummydata.columns.makeId,
          useCount: 1
        }

        if (dependency?.length) {
          obj['dependencyId'] = dependency[0]
        }

        nsOptions.push(obj)
        done(obj, 'toggle')
      }
    }
  }

  makeActuatorQuery(y = null, canShowWarningMsg = true, setItemcode = false) {
    return new Promise((resolve, reject) => {
      let actsizingData = this.actsizingdata
      if (y && _.isEmpty(this.actsizingdata?.actuator)) {
        // eslint-disable-next-line
        /* eslint-disable */
        actsizingData = this.storedummydata?.sizingdata?.['Shutoff Pressure']?.[`row_${y}`]
      }

      const curRow = actsizingData?.rowNumber
      // For Ball Valve check if Fail Close or Open and refer chart given by Mr.Gudi 90% an d80% and 50% Run
      // 1. Chekc if Group Name has  Ball Valve or TMBV or FBV and Portfolio Id Valve
      // 2. Check fail Close BTC * 90% and Fail Open BTO=90 BTC and ETC=80 and RUN =50

      const productGroupIndex = _.findIndex(this.jExcelOptions.columns, { id: 'Product_Group' })
      const productGroup = this.$refs.spreadsheet.jexcel.getValueFromCoords(productGroupIndex, curRow)
      let BTO = parseFloat(
        parseFloat(
          (actsizingData?.vtorque || 0.0000001) * (actsizingData?.SafetyFactor?.value || 0.0000001)
        ).toFixed(2)
      )

      if (this.actsizingdata?.makeId === 3) {
        BTO *= 0.113
      }
      let BTC = parseFloat(Number.parseFloat((BTO * 30) / 100).toFixed(2))
      let ETC = Number(
        Number.parseFloat(
          (actsizingData?.vtorque || 0.0000001) * (actsizingData?.SafetyFactor?.value || 0.0000001)
        ).toFixed(2)
      )

      if (this.actsizingdata?.makeId === 3) {
        ETC *= 0.113
      }

      let ETO = Number(Number.parseFloat((ETC * 30) / 100).toFixed(2))
      const style = actsizingData?.Shaft?.id || null
      let verId = (actsizingData?.actuator?.id === 184 ? 8 : actsizingData?.actuator?.id) || null
      if (verId === 193) {
        verId = 9
      }
      const supplyPressure = Number(
        Number.parseFloat(actsizingData?.SupplyPRMin?.code || 0).toFixed(2)
      )

      if (_.includes(_.toLower(productGroup), 'ball valve')) {
        if (Number(actsizingData?.Shaft?.id) === 163) {
          BTC = BTO * 0.9
          ETC = BTO * 0.8
          ETO = BTO * 0.8
        } else if (Number(actsizingData?.Shaft?.id) === 164) {
          BTC = BTO
          BTO = BTO * 0.9
          ETC = BTO * 0.8
          ETO = BTO * 0.8
        }
      }
      let productId = (
        Number(actsizingData?.actuator?.id) ||
        this.storedummydata?.sizingdata?.['Shutoff Pressure']?.[`row_${y}`]?.actuator
      )

      if (this.companydata.code === '08_DELVAL_USA') {
        if(actsizingData?.actuator?.id == 184){
          productId = 184
      }
      if(actsizingData?.actuator?.id == 185){
          productId = 10
      }
      if(actsizingData?.actuator?.id == 193){
          productId = 9
      }
    }

      this.$apollo.query({
        query: queries.getproposalpricelist,
        variables: {
          productId,
          seriesId: this.actsizingdata?.series || actsizingData?.series || null,
          orgId: !this.storedummydata.orgId
            ? Number(this.salesRegion.id)
            : this.storedummydata.orgId,
          companyId: Number(this.companydata.id),
          makeId: Number(this.storedummydata.columns.makeId.id)
        }
      }).then(({ data }) => {
        //const edges = data.getPriceList
        if (data.getPriceList.length) {
        const edges = _.find(data.getPriceList, el => el.series.name === this.actsizingdata?.series) //data.getPriceList
        // debugger
        const list = edges?.priceList
        // const list = Object.values(edges.priceList)
        this.actuatorItemcodeList = _.keys(list)
        this.actuatorprice = edges.priceList // _.fromPairs(list.map(a => _.toPairs(a)).flat()
        if (setItemcode) {
          let tempObj = {}
          _.forEach(list, a => {
            tempObj = _.merge(tempObj, a)
          })
          itemcodelist = tempObj
        }
      }
        resolve()
      }).catch(reject)

      let fetchPolicy = 'cache-first'

      if (this.oldSeriesId !== this.actsizingdata?.series) {
        fetchPolicy = 'network-only'
      }

      if (!_.includes([10, 185], productId)) {
      this.$apollo.query({
        query: queries.getActuator,
        variables: {
          btc: BTC,
          bto: BTO,
          etc: ETC,
          eto: ETO,
          verId: verId,
          style: style,
          makeId: Number(this.storedummydata.columns.makeId.id),
          supplyPressure: supplyPressure,
          seriesId: this.actsizingdata?.series || null
        },
        fetchPolicy
        // fetchPolicy: 'network-only'
      }).then(({ data: { actSizing } }) => {
        this.oldSeriesId = this.actsizingdata?.series
        if (actSizing?.length) {
          if (this.companydata?.code === '03_DELVAL') {
            this.actsizingdata = {
              ...this.actsizingdata,
              fgcode: _.map(actSizing, 'fgcode1')
                ? _.map(actSizing, 'fgcode1').toString()
                : null,
              fgDesc: _.map(actSizing, 'fgdescription1')
                ? _.map(actSizing, 'fgdescription1').toString()
                : null
            }
          } else if (this.companydata?.code === '08_DELVAL_USA') {
            this.actsizingdata['fgcode'] = _.map(actSizing, 'fgcode27')
              ? _.map(actSizing, 'fgcode27').toString()
              : null
            this.actsizingdata['fgDesc'] = _.map(
              actSizing,
              'fgdescription27'
            ) ? _.map(actSizing, 'fgdescription27').toString()
              : null
              console.log(this.actsizingdata['actuator']?.id )
              if (this.actsizingdata['actuator']?.id === 193 ){
              if (this.actsizingdata['manualOverride']?.id !== 212) {
               let MorCode = this.actsizingdata['manualOverride']?.code
               let insertCode = this.actsizingdata['insert']?.code
               let specialCode = this.actsizingdata['special']?.code

               let FGarr = this.actsizingdata['fgcode'].split('-')
               if (MorCode.length) {
                FGarr[5] = MorCode
              }

              if (insertCode?.length) {
                FGarr[7] = insertCode
              }
              if (specialCode?.length) {
                FGarr[8] = specialCode
              }

               this.actsizingdata['fgcode'] = FGarr.join('-')
              }

            }
            if (this.actsizingdata['actuator']?.id === 184) {

            let insertCode = this.actsizingdata['insert']?.code
            let specialCode = this.actsizingdata['special']?.code

            let FGarr = this.actsizingdata['fgcode'].split('-')
            if (insertCode?.length) {
              FGarr[4] = insertCode
              }
              if (specialCode?.length) {
                FGarr[5] = specialCode
              }

            this.actsizingdata['fgcode'] = FGarr.join('-')
            }
          }
        } else if (canShowWarningMsg) {
          this.showWarningMsg({
            message: 'No actuator Data for selected combination',
            textColor: 'white'
          })
        }
      })
    } else {
      const elecSizeArr = [
 {
   "id": 1,
   "size": "010",
   "code": "010",
   "torque": 885
 },
 {
   "id": 2,
   "size": "021",
   "code": "021",
   "torque": 2124
 },
 {
   "id": 3,
   "size": "050",
   "code": "050",
   "torque": 4425
 },
 {
   "id": 4,
   "size": "080",
   "code": "080",
   "torque": 7080
 },
 {
   "id": 5,
   "size": "110",
   "code": "110",
   "torque": 9735
 },
 {
   "id": 6,
   "size": "200'",
   "code": "200'",
   "torque": 17700
 },
 {
   "id": 7,
   "size": "300",
   "code": "300",
   "torque": 26550
 }
]
      const inertList = [
      {
        "id": 1,
        "code": "01D",
        "value": ".551x.394",
        "size": [1]
      },
      {
        "id": 2,
        "code": "02D",
        "value": ".630x.433",
        "size": [1,2]
      },
      {
        "id": 3,
        "code": "03D",
        "value": ".748x .512",
        "size": [1,2,3]
      },
      {
        "id": 4,
        "code": "04D",
        "value": ".866x.630",
        "size": [2,3,4]
      },
      {
        "id": 5,
        "code": "05D",
        "value": "1.18x.866",
        "size": [3,4]
      },
      {
        "id": 6,
        "code": "20D",
        "value": "1.38x.940",
        "size": [4,5,6,7]
      },
      {
        "id": 7,
        "code": "22R",
        "value": "1.38\" (.39x.39)",
        "size": [4,5,6,7]
      },
      {
        "id": 8,
        "code": "23R",
        "value": "1.97\" (.47x.39)",
        "size": [7]
      },
      {
        "id": 9,
        "code": "26R",
        "value": "1.57\" (.47x.31)",
        "size": [7]
      }
      ]
      const selAct = _.first(_.filter(elecSizeArr, x =>  x.torque >= Number(BTO/ 0.113))) // inlbs
      const size =  selAct?.size
      this.ListOfInsert = _.filter(inertList, x =>  _.includes(x.size, selAct?.id))
      this.actsizingdata['insert'] = this.ListOfInsert[0]
  // construct FGCODE and Decription and get price
      const {series, nosphaz, supplyvoltage, hertz, enclosure, insert, special} = actsizingData
      const EleActFgCode = `${series}-${size}-${nosphaz.code}-${supplyvoltage.code}-${hertz.code}-${enclosure.code}-${insert.code}-${special.code}`
      const EleActFgCDesc =  `Series :-${series}:~${series} Electric Actuator, Size :-${size}:~ ${size}, Phase:- ${hertz.code}:~${nosphaz.value}, Supply Voltage:-${supplyvoltage.code}:~${supplyvoltage.value}, Hertz:-${hertz.code}:~${hertz.value}, Enclosure:-${enclosure.code}:~${enclosure.value}, Insert:- ${insert.code}:~${insert.code}, Special:-${special.code}:~${special.value}`
      this.actsizingdata['fgcode'] = EleActFgCode
      this.actsizingdata['fgDesc'] = EleActFgCDesc
      // const eleActPrice = this.actuatorprice
      // debugger
      // console.log(this.actuatorItemcodeList, this.actuatorprice)
      }
    })
  }

  getActuator() {
    return new Promise((resolve, reject) => {
      this.$refs.sizingdata.validate()
        .then(success => {
          if (success) {
            this.makeActuatorQuery()
              .then(resolve)
          } else {
            reject()
            this.showNegativeMsg({ message: 'Select all required Fields' })
          }
        })
    })
  }

  undochanges(instance, historyRecord, allrecords) {
    const meta = instance.jexcel.options.meta
    const { action, records, columns, selection } = historyRecord || {}
    this.jExcelOptions.columns = instance.jexcel.options.columns

    if (historyRecord) {
      if (action === 'deleteColumn') {
        _.forEach(columns, col => {
          if (col.desc && this.deletedcoumn.metaData[col.title]) {
            meta[col.title] = this.deletedcoumn.metaData[col.title]
            this.deletedcoumn.metaData = _.omit(this.deletedcoumn.metaData, [
              `${col.title}`
            ])
          }
          if (this.deletedcoumn.technicaldata[col.title]) {
            this.storedummydata.technicaldata[
              col.title
            ] = this.deletedcoumn.technicaldata[col.title]
            this.deletedcoumn.technicaldata = _.omit(
              this.deletedcoumn.technicaldata,
              [`${col.title}`]
            )
          }
          if (col.sizing && this.deletedcoumn.sizingdata[col.title]) {
            this.storedummydata.sizingdata[
              col.title
            ] = this.deletedcoumn.sizingdata[col.title]
            this.deletedcoumn.sizingdata = _.omit(
              this.deletedcoumn.sizingdata,
              [`${col.title}`]
            )
          }
        })
      } else if (action === 'deleteRow') {
        instance.jexcel.options.meta = this.deletedRowdata.metaData
        this.storedummydata['sizingdata'] =
          this.deletedRowdata?.sizingdata || {}
        this.storedummydata['technicaldata'] =
          this.deletedRowdata?.technicaldata || {}
        this.deletedRowdata = {
          sizingdata: {},
          metaData: {},
          technicaldata: {}
        }
      }

      this.updateproposalRowdata(_.cloneDeep(this.storedummydata))

      if (action === 'setValue') {
        const { x, y, col, oldValue, newValue } = _.last(records)
        const cell = allrecords[y]?.[x] || null

        this.colculatePrice(instance, cell, x, y, oldValue, newValue)
      }
    }
  }

  redochange(instance, historyRecord, allrecords) {
    if (historyRecord) {
      const { action, records, columns, selection } = historyRecord || {}
      const { x, y, col, oldvalue, newValue } = _.last(records) || {}
      const cell = allrecords[y]?.[x] || null

      if (action === 'setValue') {
        this.colculatePrice(instance, cell, x, y, newValue, oldvalue)
      }
    }
  }

  colculatePrice(instance, cell, x, y, value, oldvalue) {
    const instancecolumns = instance.jexcel.options.columns
    let multipleCode = ''
    const currentCol = instancecolumns?.[Number(x)]
    const indexQty = _.findIndex(this.jExcelOptions.columns, 'qty')
    const qtyValue = parseInt(instance.jexcel.getValueFromCoords(indexQty, +y)) || 0

    if (currentCol.price && value !== undefined && value !== null) {
      value = this.unmask(value)
    }

    if (
      value?.length ||
      Number(value) > 0 ||
      currentCol.price ||
      currentCol.discount || currentCol?.type === 'dropdown'
    ) {
      const columntemplate = _.cloneDeep(this.storedummydata?.columns || [])
      const productvalue = instance.jexcel.getValueFromCoords(
        _.findIndex(this.jExcelOptions.columns, 'productcol'),
        y
      )

      if (currentCol.id && currentCol.title) {
        const curColumnName = this.jExcelOptions.columns[Number(x)]?.id

        if (
          currentCol.qty ||
          currentCol.price ||
          currentCol.unitval ||
          currentCol.testing ||
          currentCol.discount ||
          currentCol.finalval ||
          currentCol.valveprice
        ) {
          this.setfinalprice(instance, x, y, value)
        }

        // NOTE listing all fgcode related columns
        let srcValue = null
        let arrofFGCells, fgvalue, fgtype
        const validprocuct = _.find(columntemplate?.fgCodeType, prod =>
          _.includes(_.map(prod.product, 'id'), productvalue)
        )

        if (!currentCol.itemcode && currentCol.isFgColumn) {
          if (
            validprocuct &&
            productvalue?.length &&
            columntemplate?.fgCodeType?.length
          ) {
            fgtype = validprocuct
            arrofFGCells = _.map(fgtype?.columns, 'id')
            fgvalue = [{ ...fgtype, rows: [Number(y)] }]
          } else if (columntemplate?.fgCodeType?.length) {
            if (instancecolumns[x].productDependent) {
              srcValue = _.find(instancecolumns[x].source, ['id', value])
              fgtype = srcValue
                ? _.find(columntemplate?.fgCodeType, prod => {
                  return _.find(prod.product, ['id', srcValue?.Product?.[0]])
                })
                : columntemplate.fgCodeType[0]
            } else fgtype = columntemplate.fgCodeType[0]
            arrofFGCells = _.map(fgtype?.columns, 'id')
            fgvalue = [{ ...fgtype, rows: [Number(y)] }]
          }

          const derievedcolList = _.flatMap(
            columntemplate.derivedColumn,
            'columns'
          )

          if (
            arrofFGCells &&
            (_.includes(arrofFGCells, curColumnName) ||
              _.includes(_.map(derievedcolList, 'id'), curColumnName))
          ) {
            if (!_.isEmpty(instance.jexcel.options?.meta?.['fgCodeType'])) {
              const fgtypes = _.isArray(
                instance.jexcel.options.meta['fgCodeType']
              )
                ? instance.jexcel.options.meta['fgCodeType']
                : _.values(instance.jexcel.options.meta['fgCodeType'])
              const currFgTypeIndex = _.findIndex(fgtypes, ['id', fgtype?.id])
              const currentfgType = fgtypes[currFgTypeIndex]
              const oldfgTypeIndex = _.findIndex(fgtypes, x =>
                _.includes(x.rows, Number(y))
              )

              if (oldfgTypeIndex !== -1) {
                fgtypes[oldfgTypeIndex].rows.splice(
                  _.indexOf(fgtypes[oldfgTypeIndex]?.rows, Number(y)),
                  1
                )
              }

              if (currentfgType) {
                currentfgType.rows = [...currentfgType.rows, Number(y)]
                fgtypes[currFgTypeIndex] = currentfgType
              } else {
                fgtypes.push({ ...fgtype, rows: [Number(y)] })
              }
              fgvalue = fgtypes
            }
            instance.jexcel.options.meta['fgCodeType'] = fgvalue
            this.itemcodereation(
              instance,
              cell,
              x,
              y,
              value,
              currentCol,
              arrofFGCells,
              multipleCode
            )
            const productval = instance.jexcel.getValueFromCoords(
              _.findIndex(instancecolumns, 'productcol'),
              y
            )
          }
        }
      }
    }
  }

  insertcol(el, columnNumber, numOfColumns, historyRecords, insertBefore) {
    const colnum = !insertBefore
      ? Number(columnNumber) + 1
      : Number(columnNumber)
    el.jexcel.setHeader(colnum)
  }

  saveDellData(val) {
    this.toggleSpinner(true)
    const cell = this.currentcell
    this.valepopsetting = true
    if(cell?.title === 'SOV') {
      const SeriesIndex = _.findIndex(this.jExcelOptions.columns, {id: 'Series'})
      const SeriesVal = cell.instance.jexcel.getValueFromCoords(SeriesIndex, cell.pos[1]).split(':-')[1]?.trim()
      const valveFG = _.findIndex(this.jExcelOptions.columns, {id: 'Valve_P/N'})

      if( SeriesVal === 'D65') {
        const d65ValveFG = valveFG.split('-')
        // d65ValveFG[3] =
        consoel.log(d65ValveFG)
      }
    }
    // FIXME [vuex] do not mutate vuex store state outside mutation handlers
    this.storedummydata = _.cloneDeep(this.storedummydata)
    this.storedummydata['technicaldata'] = _.cloneDeep(this.getProposaldata?.technicaldata || {})

    if (val === 'save') {
      const saveval = this.formdata.multiplyval
        ? this.formdata.finalprice || 0
        : this.formdata.price || 0

      if (!this.formdata.multiplyval && this.formdata.finalprice) {
        this.formdata = _.omit(this.formdata, ['finalprice', 'multiplier'])
      }
      // console.log(this.formdata, 'this.formdata')
      cell.instance.jexcel.options.meta = {
        ...cell.instance.jexcel.options.meta,
        [cell.title]: {
          ...cell.instance.jexcel.options.meta[cell.title],
          [cell.pos[1]]: this.formdata
        }
      }
      const value = _.includes(_.toLower(this.currentcell.title), 'valve')
        ? parseFloat(saveval || 0).toFixed(this.roundCount)
        : saveval
        if (this.formdata?.desc){
      cell.instance.jexcel.setComments(cell.name, this.formdata.desc) // cell name doesnot exist for valpn
        }
      if (cell?.title?.toUpperCase() === 'VALVE P/N' || cell?.title?.toUpperCase() === 'VALVE_LP') {
        const qtyIndex =  _.findIndex(this.jExcelOptions.columns, {id: 'Qty.'})
        const vlvLpIndex = _.findIndex(this.jExcelOptions.columns, {id: 'valve_price'})
        cell.instance.jexcel.setComments(cell.name, this.formdata.desc)
        cell.instance.jexcel.setValueFromCoords(cell.pos[0], cell.pos[1], this.formdata?.itemcode)
        cell.instance.jexcel.setValueFromCoords(vlvLpIndex, cell.pos[1], this.formdata?.unitPrice)
        cell.instance.jexcel.setValueFromCoords(qtyIndex, cell.pos[1], this.formdata?.qty)
        this.cellItemcode.val = this.formdata?.itemcode
      } else {
      cell.instance.jexcel.setValueFromCoords(cell.pos[0], cell.pos[1], value)
      this.formdata = {}
      }
      this.formdata = {}
    } else {
      this.formdata = {}
      Object.keys(cell).forEach(key => {
        this.currentcell[key] = null
      })
      this.comment = false
    }

    if (cell?.title?.toUpperCase() === 'VALVE P/N') {
        this.setColValFromItemCode()
    }
    if (cell?.pos) {
    const sheet = this.$refs.spreadsheet.jexcel
      const instance = sheet.el
      const sheetoptions = instance.jexcel.options
      const instancecolumns = sheetoptions.columns
      let produtcTypeColIndex= _.filter(instancecolumns, {title:'Product Type'})?.[0]?.index
      // console.log(cell)
      let produTypeVal = instance.jexcel.getValueFromCoords(produtcTypeColIndex, cell?.pos[1])
      const productTypeCode =  produTypeVal?.split(':-')?.[1]
      if (_.includes(['x'], productTypeCode?.toLowerCase())) {
      let sysPrCode =  this.productMetaData.FGCode
      let sysPrShortDes = this.productMetaData.shortDes
      let systargetCell = ''
      const indexSysItemCode = _.findIndex(instancecolumns, 'isSystem')
      const sysColumnNameTarget = jexcel.getColumnNameFromId([indexSysItemCode, cell?.pos[1]])
      if (sysColumnNameTarget) {
        const cell = this.currentcell
        systargetCell = instance.jexcel.getCell(sysColumnNameTarget)
        systargetCell?.classList.remove('readonly')
        instance.jexcel.setValue(sysColumnNameTarget, sysPrCode)
        systargetCell?.classList.add('readonly')
        cell.instance.jexcel.options.meta = {
          ...cell.instance.jexcel.options.meta,
          ['SysItemCode']: {
            ...cell.instance.jexcel.options.meta['SysItemCode'],
            [cell.pos[1]]: { ['itemcode'] : sysPrCode, ['descLg']: sysPrShortDes }
          }
        }
      }
    } else {
     this.syscodeCreation()
    }
  }
    this.comment = false
    this.toggleSpinner(false)
  }

  deletecell(instance, cellrange) {
    const meta = _.cloneDeep(instance.jexcel.options.meta)
    const proposalobj = this.storedummydata
    for (
      let row = Number(cellrange[1]);
      row < Number(cellrange[3]) + 1;
      row++
    ) {
      for (
        let col = Number(cellrange[0]);
        col < Number(cellrange[2]) + 1;
        col++
      ) {
        const column = this.jExcelOptions.columns[col]

        if (column.sizing && proposalobj.sizingdata[column.title]) {
          if (proposalobj.sizingdata[column.title][row]) {
            delete proposalobj.sizingdata[column.title][row]
          }
        } else if (column.desc) {
          if (meta?.[column.title]?.[row]) {
            meta[column.title] = _.omit(meta[column.title], [`${row}`])
          }
          if (proposalobj?.technicaldata?.[column.title]?.[row]) {
            delete proposalobj?.technicaldata[column.title][row]
          }
        }
      }
      // jexcel.current.options.columns[columnId]
    }
    this.$set(instance.jexcel.options, 'meta', meta)
    this.updateproposalRowdata(proposalobj)
  }

  beforepaste(instance, data, x, y) {
    // this.toggleSpinner(true)
    // this.rowpaste = true
  }

  saveItemCode(type) {
    // eslint-disable-next-line no-async-promise-executor
    return new Promise(async (resolve) => {
      let orignalsrc = []
      this.toggleSpinner(true)
      this.disableitemcodesave = true
      const sheet = this.$refs.spreadsheet.jexcel
      const cols = sheet.options.columns
      const columntemplate = this.storedummydata.columns
      const stemcolumn = _.findIndex(
        cols,
        z => z.valveprice || _.includes(_.toLower(z.id), 'valve_lp')
      )
      const instance = sheet.el
      const cell = this.currentcell
      this.storedummydata = _.cloneDeep(this.storedummydata)
      const valvedesc = []
      if (type === 'save') {
        var celllist = []
        var listitemlen = []

        if (this.cellItemcode.val) {
          let fgtypeColumns = null
          const y = Number(this.cellItemcode.y)
          const fgprice = itemcodelist[this.cellItemcode.val]
          this.formdata.desc = ''
          this.formdata['price']= fgprice
          if (columntemplate?.fgCodeType?.length) {
            const fgtype = _.find(
              columntemplate?.fgCodeType,
              (prod) => {
                return _.find(
                  prod.product,
                  x => _.indexOf(this.series?.Product, x.name) > -1
                ) ? prod : false
              }
            )
            fgtypeColumns = _.filter(fgtype?.columns, z => z.coltype.id === 1)
            orignalsrc = _.map(
              _.filter(fgtype?.columns, z => z.coltype.id === 2),
              x => x?.sourceId?.name || x?.title
            )
            celllist = _.map(
              _.filter(fgtype?.columns, z => z.coltype.id === 1),
              x => x?.sourceId?.name || x?.title
            )
            let fgvalue = [{ ...fgtype, rows: [Number(y)] }]

            if (instance.jexcel.options.meta?.fgCodeType?.length) {
              const fgtypes = instance.jexcel.options?.meta?.['fgCodeType']
              const currFgTypeIndex = _.findIndex(fgtypes, ['id', fgtype?.id])
              const currentfgType = fgtypes[currFgTypeIndex]
              const oldfgTypeIndex = _.findIndex(fgtypes, x =>
                _.includes(x.rows, y)
              )

              if (oldfgTypeIndex !== -1) {
                fgtypes[oldfgTypeIndex].rows.splice(
                  _.findIndex(fgtypes[oldfgTypeIndex]?.rows, y),
                  1
                )
              }
              if (currentfgType) {
                currentfgType.rows = [...currentfgType.rows, y]
                fgtypes[currFgTypeIndex] = currentfgType
              } else {
                fgtypes.push({ ...fgtype, rows: [Number(y)] })
              }

              fgvalue = fgtypes
            }
            instance.jexcel.options.meta['fgCodeType'] = fgvalue
          } else {
            celllist = _.map(
              _.filter(columntemplate?.fgCode, z => z.type.id === 1),
              x => x.columns?.sourceId?.name || x?.columns?.title
            )
            listitemlen = _.map(
              _.filter(columntemplate?.fgCode, z => z.type.id === 1),
              x => x.code
            )
            orignalsrc = _.map(
              _.filter(columntemplate?.fgCode, z => z.type.id === 2),
              x => x.columns.id
            )
          }

          let val = null
          let itemcode = []
          let trimCodeCol = []
          const rowval = cols.map(x => '')

          cols.forEach((v, j) => {
            rowval[j] = instance.jexcel.getValueFromCoords(j, y)
          })

          if (stemcolumn !== -1) {
            rowval.splice(stemcolumn, 1, parseFloat(fgprice || 0).toFixed(this.roundCount))
          }

          if (_.includes(this.cellItemcode.val, '-')) {
            itemcode = this.cellItemcode.val.split('-')
          } else {
            const valitem = this.cellItemcode.val.split('-')
            listitemlen.forEach((item, i) => {
              const index = _.sum(listitemlen.slice(0, i))
              itemcode.push(_.sum(valitem.slice(index, index + item)))
            })
          }

          itemcode.forEach((x, i) => {
            if (
              columntemplate.derivedColumn?.length &&
              (fgtypeColumns
                ? fgtypeColumns?.[i]?.isDerived
                : _.includes(
                  _.map(columntemplate.derivedColumn, 'title'),
                  celllist[i]
                ))
            ) {
              trimCodeCol =
                columntemplate.derivedColumn[
                  _.findIndex(columntemplate.derivedColumn, [
                    'title',
                    celllist[i]
                  ])
                ]

              const trimval = _.find(
                _.cloneDeep(trimCodeCol.source),
                fg => fg.fgtCode === x || fg.fgtCode === Number(x)
              )
              trimCodeCol = trimCodeCol.columns

              if (trimval) {
                // trimval = trimval[0]
                trimCodeCol.forEach(c => {
                  val = _.find(this.jExcelOptions.columns, z => z.id === c.id)
                  val = _.filter(
                    val?.source,
                    a =>
                      a.name.split(':-')[0] ===
                      trimval?.[_.toLower(val?.sourceId.name)]
                  )
                  val = val.length ? _.first(val).name : ''
                  rowval.splice(_.findIndex(cols, ['id', c.id]), 1, val)
                })
              }
            } else {
              if (celllist[i] === 'SERIES') {
                val = _.find(cols, [
                  'sourceId.name',
                  _.upperCase(celllist[i])
                ])?.source.filter(a =>
                  _.includes(a.name, ':-')
                    ? _.split(a.name, ':-')?.[1] === x
                    : _.includes(a.name, x)
                )

                val = val?.[0]?.Product?.[0] || ''
                rowval.splice(_.findIndex(cols, 'productcol'), 1, val)

                val = _.find(cols, [
                  'sourceId',
                  {
                    id: 34,
                    name: 'END CONNECTION',
                    masterId: {
                      id: 2,
                      name: 'Product Attribute'
                    }
                  }
                ])
                if (
                  _.find(cols, [
                    'sourceId',
                    {
                      id: 34,
                      name: 'END CONNECTION',
                      masterId: {
                        id: 2,
                        name: 'Product Attribute'
                      }
                    }
                  ])
                ) {
                  val = _.find(cols, [
                    'sourceId',
                    {
                      id: 34,
                      name: 'END CONNECTION',
                      masterId: {
                        id: 2,
                        name: 'Product Attribute'
                      }
                    }
                  ])?.source.filter(a => _.includes(a.Series, x))
                  val = val.length ? val[0].name : ''
                  const endConnectionIndex = _.findIndex(cols, x =>
                    _.includes(
                      ['END CONNECTION', 'End_Connection'],
                      x?.sourceId?.name
                    )
                  )
                  if (endConnectionIndex > -1) {
                    rowval.splice(endConnectionIndex, 1, val)
                  }
                }
                if (_.find(cols, ['sourceId.name', 'LEAKAGE CLASS'])) {
                  val = _.find(cols, [
                    'sourceId.name',
                    'LEAKAGE CLASS'
                  ])?.source.filter(a => _.includes(a.Series, x))
                  val = val?.length ? val[0].name : ''
                  rowval.splice(
                    _.findIndex(cols, ['sourceId.name', 'LEAKAGE CLASS']),
                    1,
                    val
                  )
                }
              }
              let currentIndex = null
              const currentColumn = _.find(cols, (col, currindex) => {
                currentIndex = currindex
                return (
                  col?.sourceId?.name === _.upperCase(celllist[i]) ||
                  col?.sourceId?.name === celllist[i]
                )
              })

              if (currentColumn) {
                const DependentVal = currentColumn?.productDependent
                  ? [_.find(cols, 'productcol')?.title]
                  : currentColumn?.dependenciesId
                let filterSRc = _.cloneDeep(currentColumn?.source)
                const source = !DependentVal?.length
                  ? currentColumn?.source
                  : _.map(DependentVal, dep => {
                    let src = filterSRc
                    const dependentIndex = currentColumn?.productDependent
                      ? _.findIndex(cols, ['title', dep])
                      : _.findIndex(
                        cols,
                        x =>
                          _.capitalize(x?.sourceId?.name) ===
                              _.capitalize(dep?.name)
                      )
                    if (dependentIndex !== -1) {
                      const dependentSourceName = currentColumn?.productDependent
                        ? 'Product'
                        : cols[dependentIndex]?.sourceId?.name
                      const dependentval = rowval[dependentIndex]
                      src = dependentval?.length
                        ? _.filter(src, x =>
                          _.includes(dependentval, ':-')
                            ? _.includes(
                              x[_.capitalize(dependentSourceName)],
                                    _.split(dependentval, ':-')?.[1]
                            )
                            : _.includes(
                              x[_.capitalize(dependentSourceName)],
                              dependentval
                            )
                        )
                        : src
                      filterSRc = src
                    }
                    return src
                  }).flat()

                val = (
                  _.filter(source, ({ name }) => _.split(name, ':-')?.[1] === x) ||
                  _.filter(source, ({ name }) => _.includes(name, x))
                )
                val = _.first(val)?.name || ''
                rowval.splice(currentIndex, 1, val)
                valvedesc.push(`${celllist[i]} :- ${val}`)
              }
            }
          })

          _.forEach(orignalsrc, c => {
            val = _.find(cols, z => z.sourceId?.value === c || z?.id === c)?.source
            val = val?.length ? val?.[0]?.name : ''
            rowval.splice(_.findIndex(cols, ['id', c]), 1, val)
          })
          instance.jexcel.setRowData(y, rowval)
          this.disableitemcodesave = false

          if (this.companydata.id) { // NOTE old condetion if (this.companydata.id !== 1)
            const productindex = _.findIndex(
              this.jExcelOptions.columns,
              'productcol'
            )
            const currentcell = {
              title: this.jExcelOptions.columns[productindex].title,
              pos: [productindex, y]
            }
            const productname = {
              name: instance.jexcel.getValueFromCoords(productindex, y)
            }

            await this.saveTechnicalData(currentcell, productname, false)
          }
        } else {
          this.showNegativeMsg({
            type: 'warning',
            textColor: 'white',
            message: 'Please select itemcode'
          })
        }
        this.cellItemcode = { y: 0, val: '' }
      } else {
        this.cellItemcode = { y: 0, val: '' }
      }

      if (type === 'save') {
        this.formdata.desc = valvedesc.toString()
        // this.formdata.desc = 'test'
        // this.formdata['price']= fgprice
      // console.log(this.formdata, 'this.formdata')
      // setComment
      let cellname = jexcel.getColumnNameFromId([stemcolumn, cell.pos[1]])
      cell.instance.jexcel.setComments(cellname, valvedesc.toString())
      const qtyIndex =  _.findIndex(this.jExcelOptions.columns, {id: 'Qty.'})
      cell.instance.jexcel.setValueFromCoords(qtyIndex, cell.pos[1], this.formdata?.qty)
      cell.instance.jexcel.options.meta = {
        ...cell.instance.jexcel.options.meta,
        [cell.title]: {
          ...cell.instance.jexcel.options.meta[cell.title],
          [cell.pos[1]]: this.formdata
        }
      }
    } else {
      this.formdata = {}
      Object.keys(cell).forEach(key => {
        this.currentcell[key] = null
      })
    }
      this.series = null
      this.toggleSpinner(false)
      this.disableitemcodesave = false
      resolve()
    })
  }
  setValuesFromItemCode() {
    // eslint-disable-next-line no-async-promise-executor
    this.toggleSpinner(true)
    const cell = this.currentcell
    const y = Number(this.cellItemcode?.y)
    const prodmeta = this.currentcell?.instance?.jexcel?.options?.meta[this.cellItemcode.title][y]?.productMeta
    const sheet = this.$refs.spreadsheet.jexcel
    const instance = sheet.el
    const cols = sheet.options?.columns
    const fgcols1 = prodmeta?.Fgcolumns
    const fgcols = _.filter(fgcols1, (x) => (!_.includes(['Total_Amount', 'Unit_Rate_Final','Unit_Rate', 'Qty.', 'QEV','Air_Suply_PR', ], x.id)))
    const rowval = fgcols.map(x => '')
    fgcols.forEach((v, j) => {
        rowval[j] = instance.jexcel.getValueFromCoords(j, y)
      })

    _.forEach(fgcols, (c) => {
      // console.log(c)
    let val = c?.fgValue?.name !== undefined ? c?.fgValue?.name : c?.fgValue
    if (c?.multiple && _.isArray(c.fgValue)) {
        let multVal = c?.fgValue?.map(x => x.name)
        // console.log(multVal)
        let v = [';']
        multVal = [...v, ...multVal]
        multVal = multVal.join(';')
        val = multVal
        // console.log(multVal)
    }
    rowval[c.index] = val
    })
    instance.jexcel.setRowData(y, rowval)
    this.toggleSpinner(false)
    this.valepopsetting = false
}

  saveTechnicalData(column, productversion, hasid) {
    return new Promise((resolve) => {
      let queryfilter = {}

      if (hasid) {
        queryfilter = {
          makeId: { eq: this.storedummydata.columns.makeId.id },
          verId: { eq: productversion.id },
          companyId: { eq: this.companydata.id }
        }
      } else {
        queryfilter = {
          makeId: { eq: this.storedummydata.columns.makeId.id },
          ver: { name: { eq: productversion.name } },
          companyId: { eq: this.companydata.id }
        }
      }

      const variables = {
        sorting: {
          field: 'id',
          direction: 'ASC'
        },
        paging: { first: 10 },
        filter: queryfilter
      }

      this.techQueryVariable = variables

      this.$apollo.query({
        variables,
        query: masterTechData
        // fetchPolicy: 'network-only'
      }).then(({ data }) => {
        const { edges } = data.masterTeches
        const techdata = edges.length && edges[0].node
        const datasrc = techdata && _.cloneDeep(techdata.dataSrc)
        this.storedummydata = _.merge(
          _.cloneDeep(this.getProposaldata),
          _.cloneDeep(this.storedummydata)
        )

        if (!this.storedummydata.technicaldata) {
          this.storedummydata['technicaldata'] = {}
        }

        this.storedummydata.technicaldata[`row_${column.pos[1]}`] = !this
          .storedummydata.technicaldata[`row_${column.pos[1]}`]
          ? []
          : this.storedummydata.technicaldata[`row_${column.pos[1]}`]

        if (
          _.filter(
            this.storedummydata.technicaldata[`row_${column.pos[1]}`],
            x => x.column === column.title
          ).length
        ) {
          if (!this.storedummydata.technicaldata[`row_${column.pos[1]}`]) {
            this.storedummydata.technicaldata[`row_${column.pos[1]}`] = []
          }

          if (
            !_.isArray(this.storedummydata.technicaldata[`row_${column.pos[1]}`]) &&
            _.isObject(this.storedummydata.technicaldata[`row_${column.pos[1]}`])
          ) {
            this.storedummydata.technicaldata[`row_${column.pos[1]}`] = _.values(this.storedummydata.technicaldata[`row_${column.pos[1]}`])
          }

          this.storedummydata.technicaldata[`row_${column.pos[1]}`].splice(
            _.findIndex(
              this.storedummydata.technicaldata[`row_${column.pos[1]}`],
              ['column', column.title]
            ),
            1,
            {
              masterdata: techdata,
              verId: techdata?.verId,
              dataSrc: datasrc,
              column: column.title,
              portfolioId: techdata.portfolioId
            }
          )
        } else {
          if (!this.storedummydata.technicaldata[`row_${column.pos[1]}`]) {
            this.storedummydata.technicaldata[`row_${column.pos[1]}`] = []
          }

          if (
            !_.isArray(this.storedummydata.technicaldata[`row_${column.pos[1]}`]) &&
            _.isObject(this.storedummydata.technicaldata[`row_${column.pos[1]}`])
          ) {
            this.storedummydata.technicaldata[`row_${column.pos[1]}`] = _.values(this.storedummydata.technicaldata[`row_${column.pos[1]}`])
          }

          this.storedummydata.technicaldata[`row_${column.pos[1]}`].push({
            masterdata: techdata,
            verId: techdata?.verId,
            dataSrc: datasrc,
            column: column.title,
            portfolioId: techdata.portfolioId
          })
        }
        this.updateproposalRowdata(_.cloneDeep(this.storedummydata))

        resolve(data)
      })
    })
  }

  recalc(instance, x, num, history, deletedColumns) {
    // this.storedummydata = _.cloneDeep(this.getProposaldata)
    const sheet1 = this.$refs.spreadsheet.jexcel

    this.$set(sheet1.options, 'columns', this.jExcelOptions.columns)
    // const columns = _.cloneDeep(sheet1.options?.columns)
    const jsonData = _.filter(sheet1.getJson(), 'Product')
    let metadata = instance.jexcel.getMeta()
    const indexunit = _.findIndex(this.jExcelOptions.columns, 'unitval')
    const indexQty = _.findIndex(this.jExcelOptions.columns, 'qty')
    const indextotal = _.findIndex(this.jExcelOptions.columns, 'totalval')
    const indexfinal = _.findIndex(this.jExcelOptions.columns, 'finalval')
    const discount = _.findIndex(this.jExcelOptions.columns, ({ id }) => _.includes(_.toLower(id), 'discount'))

    _.forEach(deletedColumns, delcol => {
      if (delcol.price) {
        const priceindexlist = []
        this.jExcelOptions.columns.forEach((c, i) => {
          if (
            c.price &&
            !_.includes(
              [indexunit, indexQty, indextotal, indexfinal, discount],
              i
            )
          ) {
            priceindexlist.push(i)
          }
        })
        //
        // for (let index = Number(x); index < Number(x) + num; index++) {
        // x = index && _.includes(priceindexlist, Number(x))
        if (jsonData.length) {
          let outSourcedprice = 0
          jsonData.forEach((row, i) => {
            const y = i
            var valuelist = []
            priceindexlist.forEach(x => {
              let a = instance.jexcel.getValueFromCoords(x, y)
              a = _.includes(a, this.currency) ? a.split(this.currency)[1] : a
              a = isNaN(Number(a)) ? 0 : Number(a)
              // a = x === col ? isNaN(Number(value)) ? 0 : Number(value) : a
              if (a > 0) {
                outSourcedprice += a
                valuelist.push(a)
              }
            })
            // if (valuelist.length) {
            var unitcost = _.round(_.sum(valuelist), 2)
            var finalName = jexcel.getColumnNameFromId([indexfinal, y])
            var unitname = jexcel.getColumnNameFromId([indexunit, y])
            // const targetCellf = instance.jexcel.getCell(finalName)
            let discountval = parseFloat(
              instance.jexcel.getValueFromCoords(discount, y)
            )
            if (isNaN(discountval) || discountval === undefined) {
              discountval = 0
            }
            let finalunitcost = 0

            // this.company.code === '08_DELVAL_USA'

            if (this.company.code === '03_DELVAL') {
              finalunitcost = discountval
                ? (unitcost - ((unitcost - outSourcedprice) * (discountval / 100)))
                : unitcost
            } else {
              const miltiplier = discountval || 1
              finalunitcost = (outSourcedprice + ((unitcost - outSourcedprice) * miltiplier))
            }

            var columnNameTotal = jexcel.getColumnNameFromId([indextotal, y])
            let qtyvalue = parseInt(
              instance.jexcel.getValueFromCoords(indexQty, y)
            )
            if (isNaN(qtyvalue)) {
              qtyvalue = 0
            }
            const indexlum = _.findIndex(this.jExcelOptions.columns, [
              'id',
              'Additional_Charges'
            ])
            let lumpsum = parseFloat(
              instance.jexcel.getValueFromCoords(indexlum, y)
            )
            lumpsum = isNaN(lumpsum) ? 0 : lumpsum

            const totalcost =
              _.round(finalunitcost * qtyvalue, 2) +
              _.round(lumpsum * qtyvalue, 2)
            const targets = [
              [instance.jexcel.getCell(unitname), unitcost, indexunit],
              [instance.jexcel.getCell(finalName), finalunitcost, indexfinal],
              [instance.jexcel.getCell(columnNameTotal), totalcost, indextotal]
            ]

            targets.forEach((z, j) => {
              z[0].classList.remove('readonly')
              instance.jexcel.setValueFromCoords(z[2], Number(y), z[1])

              z[0].classList.add('readonly')
            })
          })
        }
        if (metadata[delcol.title]) {
          this.deletedcoumn.metaData[delcol.title] = _.cloneDeep(
            metadata[delcol.title]
          )
          metadata = _.omit(metadata, [`${delcol.title}`])
        }
      }
    })
  }

  deleterow(instance, rowNumber, numOfRows, rowRecords) {
    const rowindex = []
    const metadata = instance.jexcel.getMeta() || {}
    this.deletedRowdata = {
      metaData: metadata || {},
      sizingdata: this.storedummydata?.sizingdata || {},
      technicaldata: this.storedummydata.technicaldata || {}
    }

    for (let i = rowNumber; i < rowNumber + numOfRows; i++) {
      rowindex.push(i)
      _.forEach(this.jExcelOptions.columns, z => {
        if (z.sizing) {
          if (
            this.storedummydata?.sizingdata?.[z.title] &&
            this.storedummydata.sizingdata[z.title][i]
          ) {
            if (this.storedummydata.sizingdata[z.title][i + numOfRows]) {
              this.storedummydata.sizingdata[z.title][
                i
              ] = this.storedummydata.sizingdata[z.title][i + numOfRows]
              delete this.storedummydata.sizingdata[z.title][i + numOfRows]
            } else delete this.storedummydata.sizingdata[z.title][i]
          }
        }
        if (metadata?.[z.title]) {
          if (z.title === 'ItemCode') {
            if (metadata?.['fgCodeType']?.[i]) {
              delete metadata['fgCodeType'][i]
            }
            if (metadata?.['ItemCode']?.[`fgcode_${i}`]) {
              delete metadata?.['ItemCode']?.[`fgcode_${i}`]
            }
          } else {
            // this.deletedRowdata.metaData[z.title] = !this.deletedRowdata.metaData[z.title] ? {} : this.deletedRowdata.metaData[z.title]
            // this.deletedRowdata.metaData[z.title][i] = metadata[z.title][i]
            if (metadata?.[z.title]?.[i + numOfRows]) {
              metadata[z.title][i] = metadata[z.title][i + numOfRows]
              delete metadata[z.title][i + numOfRows]
            } else if (metadata[z.title][i]) delete metadata[z.title][i]
          }
        }
        if (z.datasheet) {
          if (
            this.storedummydata?.technicaldata?.[z.title] &&
            this.storedummydata.technicaldata[z.title][i]
          ) {
            if (this.storedummydata.technicaldata[z.title][i + numOfRows]) {
              this.storedummydata.technicaldata[z.title][
                i
              ] = this.storedummydata.technicaldata[z.title][i + numOfRows]
              delete this.storedummydata.technicaldata[z.title][i + numOfRows]
            } else delete this.storedummydata.sizingdata[z.title][i]
          }
        }
      })
    }
    _.forEach(Object.keys(metadata), m => {
      if (m === 'ItemCode') {
        const metalist = _.groupBy(
          _.toPairs(metadata[m]),
          x => _.split(x?.[0], '_')?.[0]
        )
        let newItemcodemeta = {}
        _.forEach(_.keys(metalist), c => {
          const filterKeys = _.map(rowindex, x => `${c}_${x}`)
          const filterdata = _.filter(
            metalist[c],
            x => !_.includes(filterKeys, x[0])
          )
          const objdata = _.fromPairs(
            _.map(filterdata, x => {
              if (Number(_.split(x[0], `${c}_`)[1]) > rowNumber + numOfRows) {
                x[0] = `${c}_${Number(_.split(x[0], `${c}_`)[1]) - numOfRows}`
              }
              return x
            })
          )
          newItemcodemeta = { ...newItemcodemeta, ...objdata }
        })
      } else {
        const metalist = Object.entries(metadata[m] || {})
        const nonalterdata = _.filter(
          metalist,
          a => Number(a[0]) < rowNumber + numOfRows
        )
        const alterdata = _.fromPairs(
          _.map(
            _.filter(metalist, x => Number(x[0]) > rowNumber + numOfRows),
            z => {
              z[0] = Number(z[0]) - numOfRows
              return z
            }
          )
        )
        metadata[m] = { ..._.fromPairs(nonalterdata), ...alterdata }
      }
    })
    _.forEach(_.keys(this.storedummydata?.sizingdata), s => {
      const sizinglist = Object.entries(this.storedummydata?.sizingdata?.[s])
      const nonalterdata = _.filter(
        sizinglist,
        a => Number(a[0]) < rowNumber + numOfRows
      )
      const alterdata = _.fromPairs(
        _.map(
          _.filter(sizinglist, x => Number(x[0]) > rowNumber + numOfRows),
          z => {
            z[0] = Number(z[0]) - numOfRows
            return z
          }
        )
      )
      // this.$set()
      this.storedummydata.sizingdata[s] = {
        ..._.fromPairs(nonalterdata),
        ...alterdata
      }
    })
    _.forEach(_.keys(this.storedummydata?.technicaldata), s => {
      const sizinglist = Object.entries(this.storedummydata?.technicaldata?.[s])
      const nonalterdata = _.filter(
        sizinglist,
        a => Number(a[0]) < rowNumber + numOfRows
      )
      const alterdata = _.fromPairs(
        _.map(
          _.filter(sizinglist, x => Number(x[0]) > rowNumber + numOfRows),
          z => {
            z[0] = Number(z[0]) - numOfRows
            return z
          }
        )
      )
      // this.$set()
      this.storedummydata.technicaldata[s] = {
        ..._.fromPairs(nonalterdata),
        ...alterdata
      }
    })

    // NOTE Remove deleted row data
    for (let index = rowNumber; index <= rowNumber + numOfRows; index++) {
      this.storedummydata.offerRowData = _.filter(
        this.storedummydata.offerRowData,
        (_row, i) => (i !== index)
      )

      if (metadata?.ItemCode) {
        delete metadata.ItemCode[`Trim_${index}`]
        delete metadata.ItemCode[`fgcode_${index}`]
      }
    }

    this.updateproposalRowdata(_.cloneDeep(this.storedummydata))
    instance.jexcel.options.meta = metadata

    if (!instance.jexcel.options.data.length) {
      instance.jexcel.setData([])
    }
  }

  beforedelcol(instance, x, num, history) {
    const sheet1 = instance.jexcel
    const columns = sheet1.options.columns

    for (let i = Number(x); i < Number(x) + Number(num); i++) {
      if (columns[i].fixedcol) {
        this.$q.notify({
          message: 'User not allowed to delete this column',
          type: 'negative'
        })

        return false
      }
    }

    return true
  }

  beforedeleterow(instance, x, num) {
    // const sheet1 = instance.jexcel
    // const metdata = sheet1.options.meta
    // for (let i = Number(x); i < Number(x) + Number(num); i++) {
    //   this.jExcelOptions.columns.forEach(z => {
    //     if (metdata[z.title] && metdata[z.title][i]) {
    //       delete metdata[z.title][i]
    //     }
    //   })
    // }
    // this.$set(sheet1.options, 'meta', metdata)
    // sheet1.options.meta

    return true
  }

  changeheader(el, column, oldValue, newValue, obj) {
    if (_.isNull(newValue)) return false
    const sheet1 = this.$refs.spreadsheet.jexcel
    this.$set(this.jExcelOptions, 'columns', sheet1.options.columns)
  }

  setPrice(value, _acessories) {
    if (value) {
      const row = _.find(this.itemcodeexceldata, { itemcode: value }) || {}
      const obj = {
        price: parseFloat(itemcodelist?.[value] || 0).toFixed(this.roundCount),
        desc: row?.['short Desc'] || ''
      }
      this.formdata = { ...this.formdata, ...obj }

      if (!_.isEmpty(this.formdata?.model)) {
        this.formdata['model'] = row?.Model
      }
    }
  }

  savedata(isNext = true) {
    return new Promise(resolve => {
      const storedummydata = _.cloneDeep(this.storedummydata)

      if (_.isEmpty(storedummydata['technicaldata'])) {
        storedummydata['technicaldata'] = this.getProposaldata?.technicaldata || null
      }

      const sheet = this.$refs.spreadsheet.jexcel
      this.$set(this.jExcelOptions, 'columns', sheet.options.columns)

      const omitRows = []
      const productTitle = _.find(sheet.options.columns, 'productcol')?.title

      _.forEach(sheet.getJson(), (row, index) => {
        if (_.isEmpty(row[productTitle])) {
          omitRows.push(index)
        }
      })

      const metaData = {}
      const rearrangeKeys = []
      const tempMetaData = sheet.getMeta() || {}

      _.forEach(_.keys(tempMetaData), (key) => {
        const metaKeys = _.keys(tempMetaData[key])

        _.forEach(metaKeys, (rowIndex) => {
          if (key === 'ItemCode') {
            const [itemCodeKey, ind] = _.split(rowIndex, '_')

            if (_.includes(omitRows, Number(ind))) {
              delete tempMetaData[key][rowIndex]

              if (!_.includes(rearrangeKeys, _.toLower(itemCodeKey))) {
                rearrangeKeys.push(_.toLower(itemCodeKey))
              }
            }
          } else if (key === 'fgCodeType') {
            // Fg code
          } else {
            if (_.includes(omitRows, Number(rowIndex))) {
              delete tempMetaData[key][rowIndex]

              if (!_.includes(rearrangeKeys, key)) {
                rearrangeKeys.push(key)
              }
            }
          }
        })
      })

      _.forEach(_.keys(tempMetaData), (key) => {
        const metaKeys = _.keys(tempMetaData[key])

        if (!metaData[key]) {
          metaData[key] = {}
        }

        if (key === 'fgCodeType') {
          metaData[key] = tempMetaData[key]
        }

        _.forEach(metaKeys, (rowIndex, index) => {
          if (key === 'ItemCode') {
            const [itemCodeKey] = _.split(rowIndex, '_')

            if (_.includes(_.toLower(itemCodeKey), 'trim')) {
              if (_.includes(rearrangeKeys, _.toLower(itemCodeKey))) {
                metaData[key][`Trim_${index}`] = tempMetaData[key][rowIndex]
              } else {
                metaData[key][rowIndex] = tempMetaData[key][rowIndex]
              }
            }
            if (_.includes(_.toLower(itemCodeKey), 'fgcode')) {
              if (_.includes(rearrangeKeys, key)) {
                metaData[key][`fgcode_${index}`] = tempMetaData[key][rowIndex]
              } else {
                metaData[key][rowIndex] = tempMetaData[key][rowIndex]
              }
            }
          } else if (key === 'fgCodeType') {
            // Fg code
          } else {
            if (_.includes(rearrangeKeys, key)) {
              metaData[key][index] = tempMetaData[key][rowIndex]
            } else {
              metaData[key][rowIndex] = tempMetaData[key][rowIndex]
            }
          }
        })
      })

      this.$worker.postMessage([
        JSON.stringify([
          sheet.options.columns,
          storedummydata,
          _.filter(
            sheet.getJson(),
            _.find(sheet.options.columns, 'productcol')?.title
          ),
          this.jExcelOptions,
          _.uniq(this.proposalNS),
          metaData
        ]),
        'prepare-data-for-save'
      ])
      this.$worker.onmessage = (e) => {
        if (_.last(e.data) === 'prepare-data-for-save-error') {
          const { message } = _.first(e.data)
          this.$eventBus.$emit('reset-loader')
          this.showWarningMsg({ textColor: 'white', message })
        }

        if (_.last(e.data) === 'prepare-data-for-save') {
          let showDialog = true
          const {
            jsonData,
            sheetData,
            discountamount,
            storedummydata
          } = _.first(e.data)
          this.storedummydata = _.cloneDeep(storedummydata)

          if (isNext) {
            this.currentsheetData = sheetData
          } else {
            showDialog = false
          }

          if (isNext === false) {
            // this.saveCurrentPaga()
            resolve(sheetData)
            return
          }
          if (!this.saveexceldata) {
            const self = this
            this.mounttaxdata({
              showtax: isNext,
              totalamount: sheetData.totalamount || 0,
              discountamount: discountamount || 0
            }).then((data) => {
              // this.discountprice = priceval
              this.taxes = showDialog
              // console.log(data, 'offerTax')
              this.offerTax = _.cloneDeep(data.offerTax)

              if (this.offerTax?.format?.id === 16) {
                self.discountprice = _.find(
                  this.offerTax?.taxcharges,
                  'isdiscountcol'
                )?.price
              }
              this.updateproposalRowdata({
                ...storedummydata,
                ...data.offerTax
              })

              const result = _.isEmpty(jsonData)
                ? null
                : { ...sheetData, taxData: data.offerTax }

              const totalAmountPrice = _.find(
                data.offerTax.taxcharges,
                ({ name }) => _.includes(_.toLower(name), 'total amount')
              )?.price || 0
              const totalAmountIndex = _.findIndex(
                this.offerTax.taxcharges,
                ({ name }) => _.includes(_.toLower(name), 'total amount')
              )
              this.$nextTick(() => {
                _.forEach(_.cloneDeep(data.offerTax.taxcharges), (taxRow, index) => {
                  if (_.has(taxRow, 'percent') && taxRow) {
                    this.calculatedvalue(taxRow, { template: true, index })
                  }
                })
                this.calctotal()

                this.$nextTick(() => {
                  if (this.offerTax.taxcharges[totalAmountIndex]) {
                    this.offerTax.taxcharges[totalAmountIndex].price = parseFloat(totalAmountPrice)?.toFixed(this.roundCount)
                  }
                  this.changetaxtype(this.offerTax?.taxformat)
                  resolve(result)
                })
              })
            })
          } else {
            resolve()
          }

          if (this.saveexceldata) {
            this.updateproposalRowdata(_.cloneDeep(this.storedummydata))
            const val = [true, false]
            this.$emit('downloadtemplate', val)
          }
        }
      }
    })
  }

  copycell(instance, row, string, rawdata) {
    const copiedata = []
    this.tempCopySourceData = rawdata
    const proposalobj = _.cloneDeep(this.getProposaldata)
    const metadata = instance.jexcel.getMeta()

    _.forEach(rawdata, row => {
      _.forEach(row, x => {
        const column = this.jExcelOptions.columns[x.col]

        if (column.sizing) {
          if (
            proposalobj.sizingdata &&
            proposalobj.sizingdata[column.title] &&
            proposalobj.sizingdata[column.title][`row_${x.row}`]
          ) {
            x['sizingdata'] =
              proposalobj.sizingdata[column.title][`row_${x.row}`]
            // this.$apollo.query({
            //   query: queries.getproposalpricelist,
            //   // Parameters
            //   variables: {
            //     companyId: Number(this.companydata.id),
            //     makeId: Number(proposalobj.columns.makeId.id),
            //     orgId: !this.storedummydata.orgId
            //       ? Number(this.salesRegion.id)
            //       : this.storedummydata.orgId,
            //     productId: Number(x['sizingdata'].actuator.id),
            //     seriesId: null
            //   }
            //   // fetchPolicy: 'network-only'
            // }).then(({ data }) => {
            //   const edges = data.getPriceList

            //   const list = Object.values(
            //     edges.flatMap(z => z.priceList)
            //   ).flat()

            //   this.actuatorprice = _.fromPairs(
            //     list.map(a => _.toPairs(a)).flat()
            //   )
            // })
          }
        } else if (column) {
          if (metadata[column.title] && metadata[column.title][x.row]) {
            x['meta'] = metadata[column.title][x.row]
          }
        }
        if (column.datasheet) {
          if (
            proposalobj?.technicaldata &&
            proposalobj.technicaldata[`row_${x.row}`]
          ) {
            var tech = _.find(proposalobj.technicaldata[`row_${x.row}`], [
              'column',
              column.title
            ])
            if (tech) {
              x['technicaldata'] = tech
            }
          }
        }

        copiedata.push(x)
      })
    })
    this.copiedcells = copiedata
  }

  paste(instance, data, records) {
    // console.log(this.copiedcells)

    _.forEach(records, (record, i) => {
      const { x, y, col, row, newValue, oldValue } = record
      let sourceRow = []
      let sourceMeta = []
      sourceRow = this.copiedcells?.[i]?.row
      sourceMeta = this.copiedcells?.[i]?.sourceMeta || this.copiedcells?.[i]?.meta

      // const { row: sourceRow, meta: sourceMeta } = this.copiedcells?.[i]
      const cellName = instance.jexcel.headers[col]?.innerText
      const meta = sourceMeta || { desc: null, price: newValue }

      if (
        this.copiedcells?.[i]?.['sizingdata'] &&
        this.jExcelOptions.columns[record.col]?.sizing
      ) {
        this.currentcell = {
          pos: [record.row, record.y],
          instance: instance,
          title: this.jExcelOptions.columns[record.col].title,
          name: jexcel.getColumnNameFromId([record.row, record.col])
        }
        this.storedummydata.sizingdata[this.jExcelOptions.columns[record.col].title][
          `row_${record.y}`
        ] = this.copiedcells[i]['sizingdata']
        this.setsizingdata(this.currentcell, this.copiedcells[i]['sizingdata'])

        if (this.companydata.id !== 1) {
          this.saveTechnicalData(
            this.currentcell,
            this.copiedcells[i]['sizingdata'].actuator,
            true
          )
        }
      } else {
        instance.jexcel.setMeta(cellName, String(row), meta)

        if (meta.desc) {
          instance.jexcel.setComments([x, y], meta.desc)
        }
      }

      if (
        this.storedummydata.technicaldata &&
        this.copiedcells?.[i]?.technicaldata &&
        this.jExcelOptions.columns[record.col].datasheet &&
        this.storedummydata.technicaldata[`row_${Number(record.y) - 1}`]
      ) {
        if (
          _.find(
            this.storedummydata.technicaldata[`row_${record.y}`],
            x => x.column === this.jExcelOptions.columns[record.col].title
          )
        ) {
          this.storedummydata.technicaldata[`row_${record.y}`].splice(
            _.findIndex(this.storedummydata.technicaldata[`row_${record.y}`], [
              'column',
              this.jExcelOptions.columns[record.col].title
            ]),
            1,
            this.copiedcells[i].technicaldata
          )
        } else {
          this.storedummydata.technicaldata[`row_${record.y}`] = !this.storedummydata
            .technicaldata[`row_${record.y}`]
            ? []
            : this.storedummydata.technicaldata[`row_${record.y}`]
          this.storedummydata.technicaldata[`row_${record.y}`].push(
            this.copiedcells[i].technicaldata
          )
        }
      }

      this.updateproposalRowdata(_.cloneDeep(this.storedummydata))
    })

    this.toggleSpinner(false)
    this.tempCopySourceData = []
    // this.rowpaste = false
  }

  editionstart(instance, cell, x, y, value) {
    this.rowpaste = true
    this.itemCode = null
    this.curColumn = instance.jexcel.options.columns[x]
    this.currentEditData = { instance, cell, x, y, value, column: this.curColumn }
    const columntemplate = _.cloneDeep(this.getProposaldata.columns)
    this.currentProduct = instance.jexcel.getValueFromCoords(
      _.findIndex(this.jExcelOptions.columns, 'productcol'), y
    ) || null

    if (this.currentProduct) {
      const productIndex = _.findIndex(this.jExcelOptions.columns, 'productcol')
      const productSource = instance.jexcel.options.columns[productIndex].source
      this.currentProductVersion = (
        _.find(productSource, { value: this.currentProduct }) ||
        _.find(productSource, { name: this.currentProduct })
      )
    }

    this.currentSeries = _.find(
      instance.jexcel.options.columns,
      x => (x?.sourceId?.name === 'SERIES')
    )?.sourceId || null

    // if (instance.jexcel.options.columns[x]?.id === 'Actuator') {
    //   this.makeActuatorQuery(null, false, true)
    // }

    if (instance.jexcel.options.columns[Number(x)]?.title) {
      this.currentcell = {
        instance,
        val: cell,
        pos: [x, y],
        productId: instance.jexcel.options.columns[x]?.productId?.id,
        title: instance.jexcel.options.columns[x]?.title,
        name: jexcel.getColumnNameFromId([x, y])
      }
      const metadata = _.cloneDeep(instance.jexcel.getMeta())

      this.formdata = (metadata[this.currentcell.title] && metadata[this.currentcell.title][y])
        ? metadata[this.currentcell.title][y]
        : {} //   this.formdata

        if (this.currentcell.title === 'Valve_LP' && metadata[this.currentcell.title]?.[y] === undefined) {
          this.formdata = metadata?.['Valve P/N']?.[y]
      }

      if (!_.isEmpty(this.formdata?.itemcode)) {
        const itemcode = _.isArray(this.formdata?.itemcode)
          ? _.join(this.formdata?.itemcode, '')
          : this.formdata?.itemcode
        this.itemCode = {
          id: _.join(this.formdata?.itemcode, '-'),
          code: _.replace((itemcode || ''), /-/g, '')
        }
      }
      if (instance.jexcel.options.columns[x].id === 'valve_price') {
        if (!this.formdata?.['price']) {
          const valvePriceIndex = _.findIndex(
            this.jExcelOptions.columns,
            ['id', 'valve_price']
          )
          const valvePrice = instance.jexcel.getValueFromCoords(valvePriceIndex, y)
          this.$set(
            this.formdata,
            'price',
            parseFloat(valvePrice || 0)?.toFixed(this.roundCount)
          )
        } else if (this.formdata?.['price']) {
          const valvePriceIndex = _.findIndex(
            this.jExcelOptions.columns,
            ['id', 'valve_price']
          )
          const valvePrice = instance.jexcel.getValueFromCoords(valvePriceIndex, y)
          this.$set(
            this.formdata,
            'price',
            parseFloat(this.formdata?.['price'] || 0)?.toFixed(this.roundCount)
          )
        }

        const itemcode = _.filter(
          instance.jexcel.getValueFromCoords(_.findIndex(this.jExcelOptions.columns, 'itemcode'), y)?.split('-'),
          (val, i) => (i !== 3 && 1 < 8 && val)
        )

        if (itemcode) {
          this.$set(this.formdata, 'itemcode', itemcode)
          this.itemCode = { id: itemcode?.join('-'), code: itemcode }
        }
      }

      // NOTE set outsource default value from template
      if (
        !_.has(this.formdata, 'outSource') &&
        _.has(instance.jexcel.options.columns[x], 'outSource')
      ) {
        this.$set(this.formdata, 'outSource', instance.jexcel.options.columns[x].outSource)
      }

      const productvalue = instance.jexcel.getValueFromCoords(
        _.findIndex(this.jExcelOptions.columns, 'productcol'),
        y
      )
      const arrofFGCells = columntemplate?.fgCode
      const derivedColumn = columntemplate?.derivedColumn
      this.currentcell = {
        instance,
        val: cell,
        pos: [x, y],
        productId: instance.jexcel.options.columns[x]?.productId?.id,
        title: instance.jexcel.options.columns[x]?.title,
        name: jexcel.getColumnNameFromId([x, y])
      }

      if (
        _.includes(
          _.map(arrofFGCells, 'columns.title'),
          instance.jexcel.options.columns[Number(x)].title
        ) &&
        _.find(
          arrofFGCells,
          { columns: { title: instance.jexcel.options.columns[Number(x)].title } }
        )?.type?.id === 1
      ) {
        if (
          instance.jexcel.getValueFromCoords(
            _.findIndex(this.jExcelOptions.columns, [
              'sourceId.name',
              'SERIES'
            ]),
            y
          )?.length
        ) {
          let seriesIndex
          const series = _.find(instance.jexcel.options.columns, (x, i) => {
            if (x?.sourceId?.name === 'SERIES') {
              seriesIndex = i
              return x
            }
          })
          const priceList = _.cloneDeep(this.getPriceList)
          this.currentPriceList = _.find(priceList, {
            companyId: this.companydata?.id,
            series: { name: series?.name }
          })
          if (
            this.currentPriceList &&
            _.includes(this.currentPriceList.orgId, this.storedummydata?.orgId)
          ) {
            itemcodelist = this.currentPriceList.priceList
            // this.itemcodeexceldata = this.currentPriceList.excelSheetData
          } else {
            this.getitemcodes(columntemplate.makeId, {
              field: 'series',
              value: {
                id: instance.jexcel.getValueFromCoords(seriesIndex, y),
                name: instance.jexcel.getValueFromCoords(seriesIndex, y)
              }
            })
          }
        }
      } else if (
        derivedColumn?.length &&
        _.includes(
          _.map(derivedColumn, z => _.map(z.columns, x => x.title)).flat(),
          instance.jexcel.options.columns[Number(x)].title
        )
      ) {
        _.forEach(derivedColumn, z => {
          if (
            _.includes(
              _.map(z.columns, x => x.title),
              instance.jexcel.options.columns[Number(x)].title
            )
          ) {
            if (
              _.includes(
                arrofFGCells.map(x => x.columns?.title),
                z.title
              ) &&
              _.find(arrofFGCells, ['columns.title', z.title]).type.id === 1
            ) {
              this.getitemcodes(columntemplate.makeId, {
                field: 'series',
                value: {
                  id: instance.jexcel.getValueFromCoords(
                    _.findIndex(this.jExcelOptions.columns, [
                      'sourceId.id',
                      31
                    ]),
                    y
                  ),
                  name: instance.jexcel.getValueFromCoords(
                    _.findIndex(this.jExcelOptions.columns, [
                      'sourceId.id',
                      31
                    ]),
                    y
                  )
                }
              })
            }
          }
        })
      }
      if (this.curColumn.productId?.id !== this.productMetaData?.productId) {
        let meta = instance.jexcel.getMeta()
        if (_.has(meta, this.curColumn.title)) {
          this.productMetaData = meta[this.curColumn.title][y]?.productMeta
        } else {
          this.productMetaData = {}
        }
      } else {
        this.productMetaData = this.formdata['productMeta']
      }
      if (
        !_.includes(
          _.toLower(instance.jexcel.options.columns[Number(x)]?.id),
          'valve_lp'
        )
      ) {
        if (this.jExcelOptions.columns[Number(x)]?.desc) {
          itemcodelist = []
          this.modellist = []
          this.listOfItemcode = []
          this.currentcell.val = cell
          this.currentcell['datasheet'] = this.jExcelOptions.columns[
            Number(x)
          ].datasheet
          if (this.currentcell['datasheet']) {
            this.currentcell['productId'] = this.jExcelOptions.columns[
              Number(x)
            ].productId
          }
          this.currentcell.pos = [x, y]
          this.currentcell['name'] = jexcel.getColumnNameFromId([x, y])
          this.currentcell['instance'] = instance
          this.currentcell['title'] = this.jExcelOptions.columns[
            Number(x)
          ].title
          const metadata = _.cloneDeep(instance.jexcel.getMeta())
          this.formdata = this.company.code === '08_DELVAL_USA'
            ? {
              ...this.formdata,
              multiplyval: true,
              multiplier: (!this.formdata.multiplier || this.formdata.multiplier < 1)
                ? 1
                : this.formdata.multiplier
            }
            : this.formdata

          if (!this.currentcell.productId) {
            this.formdata = (metadata[this.currentcell.title] && metadata[this.currentcell.title][y])
              ? metadata[this.currentcell.title][y]
              : this.formdata
          }
          // console.log(this.currentcell)
          if (this.currentcell.productId) {
            // check if priceList exist then check from store else do an api Call // removed
            const priceList = _.cloneDeep(this.getPriceList)
            this.currentPriceList = _.find(priceList, {
              companyId: this.companydata?.id,
              productId: {
                id: this.currentcell.productId.id,
                name: this.currentcell.productId.name
              }
            })
            if (
              _.includes(
                this.currentPriceList?.orgId,
                this.storedummydata?.orgId
              )
            ) {
              itemcodelist = this.currentPriceList.priceList
              this.itemcodeexceldata = _.cloneDeep(this.currentPriceList.excelSheetData)
            } else {
              itemcodelist = []
              this.itemcodeexceldata = []
              this.getitemcodes(
                columntemplate.makeId,
                {
                  field: 'productId',
                  value: this.jExcelOptions.columns[Number(x)]?.productId
                },
                true
              )
            }
          }
          // console.log(this.editId, this.currentcell.productId)
          if (this.currentcell.productId || this.curColumn.productId) {
            this.editId = this.currentcell.productId?.id || this.curColumn.productId?.id
          }
          this.comment = true
        } else if (
          Number(x) === _.findIndex(this.jExcelOptions.columns, 'itemcode')
        ) {
          this.cellItemcode['y'] = y
          this.cellItemcode['val'] = cell.innerText
          this.cellItemcode['title'] = this.jExcelOptions.columns[
            Number(x)
          ].title
          if (this.company?.code  === '71') {
            this.isCodeBuilder = true
            this.comment = true
            this.editId = 1
            this.lineproductId = productvalue
          } else {
            this.showseries = true
          }
          // this.code = true
        } else if (this.jExcelOptions.columns[Number(x)]?.sizing) {
          this.currentcell.val = cell
          this.currentcell.pos = [x, y]
          this.currentcell['name'] = jexcel.getColumnNameFromId([x, y])
          this.currentcell['instance'] = instance
          this.currentcell['title'] = this.jExcelOptions.columns[
            Number(x)
          ].title

          if (
            this.storedummydata.sizingdata &&
            this.storedummydata.sizingdata[
              this.jExcelOptions.columns[Number(x)].title
            ] &&
            this.storedummydata.sizingdata[
              this.jExcelOptions.columns[Number(x)].title
            ][`row_${y}`]
          ) {
            this.actsizingdata = _.cloneDeep(
              this.storedummydata.sizingdata[
                this.jExcelOptions.columns[Number(x)].title
              ][`row_${y}`]
            )
          } else {
            this.setActSizingDefaultData()
          }
          if (productvalue?.length) {
            if (_.isEmpty(this.actsizingdata)) {
              this.actsizingdata = {}
            }

            this.actsizingdata['valveType'] = {
              name: productvalue
            }
          }

          this.actsizingdata['rowNumber'] = Number(y)
          let celltorque =  Number(
              instance.jexcel.getValueFromCoords(
                _.findIndex(this.jExcelOptions.columns, ['id', 'Valve_Torq']),
                y
              )
            )
            if (celltorque) {
              this.actsizingdata['vtorque'] = celltorque
            }
          this.actsizingdata = {
            ...this.actsizingdata
          }
          if (this.companydata.id !== 2) {
            this.actsizingdata['Media'] = instance.jexcel.getValueFromCoords(
              _.findIndex(this.jExcelOptions.columns, ['id', 'Service']),
              y
            )
          }

          this.actsizing = true
        }

        if (instance.jexcel.options.columns[Number(x)]?.id === 'Actuator_price') {
          if (_.isEmpty(this.itemcodeexceldata)) {
            _.forEach(this.actuatorItemcodeList, (code) => {
              this.itemcodeexceldata.push({
                itemcode: code,
                Model: this.formdata.model
              })
            })
          }
        }

        if (
          this.jExcelOptions.columns[Number(x)]?.id !== 'Series' &&
          this.jExcelOptions.columns[Number(x)]?.id !== 'Valve_Size'
        ) {
          const indexseries = _.findIndex(this.jExcelOptions.columns, [
            'id',
            'Series'
          ])
          const seriesval = instance.jexcel.getValueFromCoords(indexseries, y)
          const indexsize = _.findIndex(this.jExcelOptions.columns, [
            'sourceId.name',
            'SIZE'
          ])
          const sizeval = instance.jexcel.getValueFromCoords(indexsize, y)
          const indexqty = _.findIndex(this.jExcelOptions.columns, 'qty')
          const qtyval = instance.jexcel.getValueFromCoords(indexqty, y)

          this.$emit('series', {
            series: seriesval,
            size: sizeval,
            Qty: qtyval
          })
        }
      }

      this.$nextTick(() => {
        if (
          _.includes(
            ['Actuator'],
            instance.jexcel.options.columns[Number(x)].title
          )
        ) {
          this.currentcell['productId'] = this.storedummydata?.sizingdata?.['Shutoff Pressure']?.[`row_${y}`]?.actuator

          if (_.isEmpty(this.itemcodeexceldata)) {
            if (_.isEmpty(this.actsizingdata?.actuator)) {
              this.actsizingdata = this.storedummydata?.sizingdata?.['Shutoff Pressure']?.[`row_${y}`]
            }

            this.getitemcodes(
              this.storedummydata.columns.makeId,
              {
                field: 'productId',
                value: this.jExcelOptions.columns[Number(x)]?.productId
              },
              true
            )
          }
        }
      })

      if (_.isArray(this.itemCode?.code)) {
        this.itemCode.code = _.join(this.itemCode.code, '')
      }
    } else {
      this.$q.notify({
        message: 'please rename the column first',
        type: 'info'
      })
    }
  }

  saveTemplateSourceData() {
    return this.saveUserTemplate({
      id: this.getProposaldata.columns.id,
      columns: this.jExcelOptions.columns
    }, this.user)
  }

  handleProductSavedEvent({ data, resultName }) {
    const response = data[resultName]
    const index = _.findIndex(this.jExcelOptions.columns, 'productcol')
    const lastIndex = this.jExcelOptions.columns[index].source.length - 1
    const addNewSource = _.cloneDeep(this.jExcelOptions.columns[index].source[lastIndex])
    this.jExcelOptions.columns[index].source[lastIndex] = {
      _id: response.id,
      id: response.name,
      name: response.name,
      'Product Group': response?.group?.name || ''
    }
    this.jExcelOptions.columns[index].source.push({
      ...addNewSource,
      id: this.jExcelOptions.columns[index].source.length
    })
    this.jExcelOptions.columns[index]['sourceorignal'] = this.jExcelOptions.columns[index].source

    this.$nextTick(() => {
      this.saveTemplateSourceData()
        .then(() => {
          this.addNewOptionDialog = false

          if (!_.isEmpty(this.currentEditData)) {
            const { instance, cell, x, y, value } = this.currentEditData
            instance.jexcel.setValueFromCoords(
              x,
              y,
              response.name,
              true
            )
          }
        }).catch(() => {
          this.addNewOptionDialog = false
        })
    })
  }

  /**
   * @method beforeChange
   * @event onbeforechange
   * @param { Object } instance
   * @param { Object } cell
   * @param { number } x
   * @param { number } y
   * @param { number|text|object|array } vallue
   * @return {boolean} if to proceed to change event
   * @document  trigger on before changed event
   */
  beforeChange(instance, cell, x, y, value) {
    this.rowPaste = true
    const proposaldata = _.cloneDeep(this.getProposaldata)
    const column = instance.jexcel.options.columns[x]
    const skipNegativeColumns = ['Discount', 'Discount %']

    // if (column?.type === 'numeric') {
    //   // eslint-disable-next-line no-useless-escape
    //   // console.log(value.replace(/[^0-9\.]/g, ''))
    // }

    // // NOTE Number field validation
    // if (
    //   column?.type === 'numeric' &&
    //   // eslint-disable-next-line no-useless-escape
    //   Number(value) !== Number(value.replace(/[^0-9\.]/g, ''))
    // ) {
    //   // eslint-disable-next-line no-useless-escape
    //   instance.jexcel.setValueFromCoords(x, y, value.replace(/[^0-9\.]/g, ''))
    //   return 0
    // }

    if (value === 'add-new' && this.curColumn?.productcol) {
      this.addNewOptionDialog = true
      return
    }

    if (
      !column.negetive &&
      column.type === 'numeric' &&
      !_.includes(skipNegativeColumns, column.id)
    ) {
      if (Number(value) < 0) {
        this.$q.notify({
          message: 'add positive value for ' + column.title,
          type: 'negative'
        })
        return Math.abs(!Number(value) ? 0 : Number(value))
      }
    }

    if (value === 'NS:-NS' || value === 'Special Requirement:-S') {
      cell.classList.remove('dropdown')
      cell.style.cssText =
        'text-align: center; background-color: rgb(255, 245, 157);'
      nsOptions = []
      this.nonStandardList = []
      this.nsMdl = true
      this.nsColumnNameTarget = column
      cell.innerText = this.nsText
      nsDependentopt = column.dependenciesId?.length
        ? _.find(instance.jexcel.options.columns, [
          'sourceId.id',
          column.dependenciesId[0].id
        ])?.source
        : []
    } else {
      cell.style.cssText =
        'text-align: center; background-color: rgb(255, 255, 255);'
    }

    if (value === '') {
      if (
        proposaldata.technicaldata &&
        proposaldata.technicaldata[column.title] &&
        proposaldata.technicaldata[column.title][y]
      ) {
        delete proposaldata.technicaldata[column.title][y]
        this.updateproposalRowdata(proposaldata)
      }
      const propmeta = instance.jexcel.getMeta()
      if (propmeta[column.title] && propmeta[column.title][y]) {
        delete propmeta[column.title][y]
        instance.jexcel.options.meta[column.title] = propmeta[column.title]
      }
      const cellname = jexcel.getColumnNameFromId([x, y])
      if (instance.jexcel.getComments(cellname)) {
        instance.jexcel.setComments(cellname, '')
      }
    }

    // return Number(value)
  }

  saveNs() {
    if (this.$refs.nstext.validate()) {
      const celldata = _.cloneDeep(this.currentcell)
      const nsObj = {
        id: `${this.nsText?.name || ''} :-NS`,
        name: `${this.nsText?.name || ''} :-NS`
      }

      if (this.nsText?.dependencyId?.name) {
        nsObj[_.capitalize(this.nsText?.dependencyId?.name)] = _.map(
          this.nsDependentvalue,
          x => _.split(x.id, ':-')[1]
        )
      }

      const targetcolumn = _.findIndex(this.jExcelOptions.columns, [
        'title',
        this.nsColumnName?.title
      ])

      if (targetcolumn !== -1) {
        celldata.instance.jexcel.options.columns[targetcolumn].source.push(
          nsObj
        )
        if (
          !_.find(this.jExcelOptions.columns[targetcolumn].sourceorignal, nsObj)
        ) {
          this.jExcelOptions.columns[targetcolumn].source.push(nsObj)

          if (!this.jExcelOptions.columns[targetcolumn]?.sourceorignal) {
            this.jExcelOptions.columns[targetcolumn]['sourceorignal'] = this.jExcelOptions.columns[targetcolumn].source
          } else {
            this.jExcelOptions.columns[targetcolumn].sourceorignal.push(nsObj)
          }
        }
        celldata.instance.jexcel.setValueFromCoords(
          celldata.pos[0],
          celldata.pos[1],
          nsObj?.name
        )
      }

      if (!this.nsText?.id) {
        let obj = this.nsText

        if (this.nsDependentvalue?.length) {
          obj = { ...obj, dependentvalues: this.nsDependentvalue }
        }

        this.saveNonStandardData(_.cloneDeep(obj))
          .then(({ data }) => {
            if (data?.nonStandard?.id) {
              this.proposalNS.push(data.nonStandard.id)
            }
          })
        this.nsText = null
        // if (obj?.id) {
        // }
      } else if (this.nsText?.id && this.nsDependentvalue?.length) {
        const obj = {
          id: this.nsText?.id,
          dependentvalues: this.nsDependentvalue,
          useCount: this.nsText.useCount + 1
        }
        this.saveNonStandardData(obj).then(({ data }) => {
          if (data?.nonStandard?.id) {
            this.proposalNS.push(data.nonStandard.id)
          }
          // this.showPositiveMsg({ message: 'Column Added successfully!' })
        })
      } else if (this.nsText?.id) {
        this.proposalNS.push(this.nsText?.id)
      }

      this.nsText = null
      this.nsDependentvalue = null
      this.nsMdl = false
      nsOptions = []
      this.nonStandardList = []
    } else {
      this.showNegativeMsg({ message: 'Value is required' })
    }
  }

  newRow() {
    const sheet = this.$refs.spreadsheet.jexcel
    sheet.insertRow()
  }

  // selectionActive(instance, x1, y1, x2, y2, origin) {},
  dropdownFilter(instance, cell, c, r, source) {
    const columnList = instance.jexcel.options.columns
    const columnData = columnList[c]
    const DependentVal = _.keys(
      _.omit(columnData?.sourceorignal[0], [
        'id',
        'name',
        '_id',
        'text',
        'value'
      ])
    )

    if (columnData?.productcol) {
      _.forEach(DependentVal, dep => {
        const dependentIndex = _.findIndex(columnList, ['title', dep])
        if (dependentIndex !== -1) {
          const dependentval = instance.jexcel.getValueFromCoords(
            dependentIndex,
            r
          )
          const orignalsrc = columnData?.sourceorignal
          source = dependentval?.length
            ? _.filter(orignalsrc, x => (
              _.includes(x['id'], 'add-new') ||
              _.trim(x[dep]) === _.trim(dependentval)
            )) : orignalsrc
        }
      })
    } else {
      source = columnData?.sourceorignal
      _.forEach(DependentVal, dep => {
        const dependentIndex = columnData?.productDependent
          ? _.findIndex(columnList, ['title', dep])
          : _.findIndex(
            columnList,
            x => _.capitalize(x?.sourceId?.name) === dep
          )

        if (dependentIndex !== -1) {
          const dependentSourceName = columnData?.productDependent
            ? 'Product'
            : columnList[dependentIndex]?.sourceId?.name
          const dependentval = instance.jexcel.getValueFromCoords(
            dependentIndex,
            r
          )
          source = dependentval?.length
            ? _.filter(source, x => (
              (_.includes(dependentval, ':-')
                ? _.includes(x[_.capitalize(dependentSourceName)], _.split(dependentval, ':-')?.[1])
                : _.includes(x[_.capitalize(dependentSourceName)], dependentval)) ||
              _.includes(x['id'], 'add-new')
            ))
            : source
        }
      })
    }

    const seriesValue = instance.jexcel.getValueFromCoords(
      _.findIndex(columnList, ['sourceId.name', 'SERIES']),
      r
    )

    if (
      _.includes(seriesValue, ':-OS') ||
      _.includes(seriesValue, ':-NS') ||
      _.includes(seriesValue, 'NS:-NS') ||
      _.includes(seriesValue, 'OS:-OS') ||
      instance.jexcel.options.columns?.[c].isNS
    ) {
      const productvalue = instance.jexcel.getValueFromCoords(
        _.findIndex(this.jExcelOptions.columns, 'productcol'),
        r
      )

      if (!source.length && productvalue) {
        source = _.uniqBy([
          { id: 'NS:-NS', name: 'NS:-NS' },
          ...(source.length ? source : columnData?.sourceorignal)
        ], 'id')
      } else if (source?.length && productvalue) {
        source = _.uniqBy(source, 'id')
      } else {
        source = _.uniqBy([
          ...(source.length ? source : columnData?.sourceorignal),
          { id: 'NS:-NS', name: 'NS:-NS' }
        ], 'id')
      }
    }

    return source
  }

  // setOption
  saveSizing(type) {
    if (type === 'save' || type === 'update') {
      this.storedummydata = _.merge(
        _.cloneDeep(this.getProposaldata),
        _.cloneDeep(this.storedummydata)
      )
      if (!this.storedummydata['sizingdata']) {
        this.storedummydata['sizingdata'] = {}
      }
      const cell = this.currentcell
      const sizedata = this.actsizingdata
      const clearActSize = (type === 'save')
      this.setsizingdata(this.currentcell, this.actsizingdata, clearActSize)

      if (this.companydata.id !== 1) {
        this.saveTechnicalData(cell, sizedata.actuator, true)
      }
      const saveval = sizedata.diffPr
      cell.instance.jexcel.setValueFromCoords(cell.pos[0], cell.pos[1], saveval)

      if (type === 'save') {
        this.setActSizingDefaultData()
      }
    } else if (type === 'close') {
      this.setActSizingDefaultData()
    }
  }

  setsizingdata(cell, sizedata, clearActSize = true) {
    if (!sizedata) return
    const fgdesc = _.split(sizedata.fgDesc, ',')
    const acttype = sizedata.actuator?.name
    const actmodel = _.find(fgdesc, z => _.includes(z, 'CYLINDER SIZE'))
    const torqueColumnIndex = _.findIndex(
      this.jExcelOptions.columns,
      x => x.id === 'Valve_Torq' || x.id === 'Valve_Torque'
    )
    if (sizedata.vtorque && torqueColumnIndex > -1) {
      cell.instance.jexcel.setValueFromCoords(
        torqueColumnIndex,
        cell.pos[1],
        sizedata.vtorque
      )
    }
    if (sizedata.Media && this.companydata.id !== 2) {
      cell.instance.jexcel.setValueFromCoords(
        _.findIndex(this.jExcelOptions.columns, ['id', 'Service']) !== -1
          ? _.findIndex(this.jExcelOptions.columns, ['id', 'Service'])
          : _.findIndex(this.jExcelOptions.columns, ['id', 'Media']),
        cell.pos[1],
        sizedata.Media
      )
    }
    if (_.findIndex(this.jExcelOptions.columns, ['id', 'Act_Type']) > -1) {
      cell.instance.jexcel.setValueFromCoords(
        _.findIndex(this.jExcelOptions.columns, ['id', 'Act_Type']),
        cell.pos[1],
        acttype
      )
    }
    const actmodelindex = _.findIndex(
      this.jExcelOptions.columns,
      x => x.id === 'Act_Model' || x.id === 'Actuator_Model'
    )
    // debugger
    if (this.storedummydata.columns?.makeId?.id === 1) {
      const modelval = [
        _.split(
          _.find(fgdesc, z => _.includes(z, 'MODEL')),
          '~'
        )?.[1]?.trim() === '--'
          ? ''
          : _.split(
            _.find(fgdesc, z => _.includes(z, 'MODEL')),
            '~'
          )?.[1]?.trim(),
        _.split(actmodel, '~')?.[1]?.trim(),
        _.split(
          _.find(fgdesc, z => _.includes(z, 'ACTUATOR TYPE')),
          '~'
        )[1]?.trim(),
        _.find(fgdesc, z => _.includes(z, 'NO. OF SPRING'))?.split('~')?.[1]?.trim()
      ]
      cell.instance.jexcel.setValueFromCoords(
        actmodelindex,
        cell.pos[1],
        modelval.filter(z => z !== '' && z !== '-').join('-')
      )
      cell.instance.jexcel.setValueFromCoords(
        _.findIndex(this.jExcelOptions.columns, ['id', 'Fail_Position']),
        cell.pos[1],
        `${
          _.find(fgdesc, z => _.includes(z, 'SHAFT ROTATION'))?.split('~')[1]
        }`
      )
      const minairprindex = _.findIndex(
        this.jExcelOptions.columns,
        x => x.id === 'Sized_@Min.Air_Supply' || x.id === 'Air_Suply_PR'
      )
      if (minairprindex > -1) {
        cell.instance.jexcel.setValueFromCoords(
          minairprindex,
          cell.pos[1],
          sizedata.SupplyPRMin.value
        )
      }
    } else if (this.storedummydata.columns?.makeId?.id === 27) {
        const minairprindex = _.findIndex(
            this.jExcelOptions.columns,
            x => x.id === 'Sized_@Min.Air_Supply' || x.id === 'Air_Suply_PR'
          )
          const supplVoltIndex = _.findIndex(
            this.jExcelOptions.columns,
            x => x.id === 'Supply_Voltage'
          )
        if (sizedata.actuator.id !== 185){
          if (minairprindex > -1) {
            cell.instance.jexcel.setValueFromCoords(
              minairprindex,
              cell.pos[1],
              'sizedata.SupplyPRMin.value'
            )
          }
          if (supplVoltIndex > -1) {
            cell.instance.jexcel.setValueFromCoords(
              minairprindex,
              cell.pos[1],
              '-'
            )
          }
          if (
            _.findIndex(this.jExcelOptions.columns, ['id', 'Fail_Position']) > -1
          ) {
            cell.instance.jexcel.setValueFromCoords(
              _.findIndex(this.jExcelOptions.columns, ['id', 'Fail_Position']),
              cell.pos[1],
              `${
                _.find(fgdesc, z => _.includes(z, 'SHAFT ROTATION'))?.split(
                  '~'
                )[1]
              }`
            )
          }

          cell.instance.jexcel.setValueFromCoords(
            _.findIndex(this.jExcelOptions.columns, [
              'id',
              'Sized_@Min.Air_Supply'
            ]),
            cell.pos[1],
            sizedata.SupplyPRMin.value
          )
          if (_.includes([8, 184],sizedata.actuator?.id)) {
            if (actmodelindex > -1) {
              cell.instance.jexcel.setValueFromCoords(
                actmodelindex,
                cell.pos[1],
                `${_.find(fgdesc, z => _.includes(z, 'SIZE'))?.split('~')[1]}-${
                  _.find(fgdesc, z => _.includes(z, 'NUMBER OF SPRING'))?.split(
                    '~'
                  )[1]
                }`
              )
            }
          } else if (_.includes([9, 193],sizedata.actuator?.id)) {
            if (actmodelindex) {
              cell.instance.jexcel.setValueFromCoords(
                actmodelindex,
                cell.pos[1],
                `${
                  _.find(fgdesc, z => _.includes(z, 'MODEL'))?.split('~')[1]
                }-${
                  _.find(fgdesc, z => _.includes(z, 'CYLINDER SIZE'))?.split(
                    '~'
                  )[1]
                }-${
                  _.find(fgdesc, z =>
                   ( _.includes(z, 'TYPE') ||  _.includes(z, 'TYPE'))
                  )?.split(':-')[1].split('~')[0]
                }-${
                  _.find(fgdesc, z =>
                   ( _.includes(z, 'SPRING PRESSURE RATING'))
                  )?.split('~')[1]
                }`
              )
            }
          }
        } else {
          const actmodelindex = _.findIndex(
      this.jExcelOptions.columns,
      x => x.id === 'Act_Model' || x.id === 'Actuator_Model'
    )
    if (actmodelindex > -1) {
        cell.instance.jexcel.setValueFromCoords(
          actmodelindex,
          cell.pos[1],
          sizedata.fgcode
        )
      }
        if (minairprindex > -1) {
          cell.instance.jexcel.setValueFromCoords(
            minairprindex,
            cell.pos[1],
            '-'
          )
        }
        if (supplVoltIndex > -1) {
          cell.instance.jexcel.setValueFromCoords(
            supplVoltIndex,
            cell.pos[1],
            sizedata.supplyvoltage.value
          )
      }
     }
    }

    let model = ''
    const fgCode = sizedata?.fgCode || sizedata?.fgcode
    let actModelIndex = _.findIndex(this.jExcelOptions.columns, ['title', 'Act_Model'])
    if (actModelIndex === -1) {
      actModelIndex = _.findIndex(this.jExcelOptions.columns, ['title', 'Actuator_Model'])
    }

    if (actModelIndex > -1) {
      model = cell.instance.jexcel.getValueFromCoords(actModelIndex, cell.pos[1])
    }
    const formdata = {
      model,
      price: 0,
      qty: 1,
      itemcode: fgCode,
      desc: sizedata.fgDesc // _.map(sizedata.fgDesc.split(','), x => x.length ? _.split(x, ':~')[0] : '').join(),
    }

    if (this.actuatorprice && Object.keys(this.actuatorprice).length) {
      const itemcode = Object.keys(this.actuatorprice)
      let sizingitemcode = fgCode
      if (_.includes([184],sizedata.actuator?.id)) {
       const newFgCode = fgCode.split('-')
       newFgCode[2] ='12'
       sizingitemcode = newFgCode.join('-')
      }
      // const sizingitemcode = fgCode.slice(0, itemcode[0]?.length)
      // formdata['price'] = Number(this.actuatorprice[sizingitemcode])
      if (_.includes(itemcode, sizingitemcode)) {
        formdata['price'] = Number(this.actuatorprice[sizingitemcode])
      }
    } else if (
      Number(
        cell.instance.jexcel.getValueFromCoords(
          _.findIndex(this.jExcelOptions.columns, ['title', 'Actuator']),
          cell.pos[1]
        )
      ) > 0
    ) {
      formdata['price'] = Number(
        cell.instance.jexcel.getValueFromCoords(
          _.findIndex(this.jExcelOptions.columns, ['title', 'Actuator']),
          cell.pos[1]
        )
      )
    }
    // actuatorprice
    // debugger
    // if (formdata['price'] > 0) {
      cell.instance.jexcel.setValueFromCoords(
        _.findIndex(this.jExcelOptions.columns, ['title', 'Actuator']),
        cell.pos[1],
        formdata.price
      )
    // }
    cell.instance.jexcel.options.meta = {
      ...cell.instance.jexcel.options.meta,
      Actuator: {
        ...cell.instance.jexcel.options.meta['Actuator'],
        [cell.pos[1]]: formdata
      }
    }
    const getColumnNameFromId = (cell.instance?.jexcel?.getColumnNameFromId instanceof Function)
      ? cell.instance?.jexcel?.getColumnNameFromId
      : jexcel?.getColumnNameFromId
    const cellname = getColumnNameFromId([
      _.findIndex(this.jExcelOptions.columns, ['title', 'Actuator']),
      cell.pos[1]
    ])
    if (cellname) {
      cell.instance.jexcel.setComments(cellname, sizedata.fgDesc)
    }
    // cellname = cell.instance?.jexcel?.getColumnNameFromId([_.findIndex(this.jExcelOptions.columns, ['title', 'Actuator']), cell.pos[1]])
    // cell.instance.jexcel?.setComments(cellname, formdata)
    this.storedummydata = _.merge(
      _.cloneDeep(this.getProposaldata),
      _.cloneDeep(this.storedummydata)
    )
    if (!this.storedummydata['sizingdata']) {
      this.storedummydata['sizingdata'] = {}
    }
    const storedata = this.storedummydata
    if (!storedata.sizingdata[cell.title]) {
      storedata.sizingdata[cell.title] = { [`row_${cell.pos[1]}`]: sizedata }
    } else {
      storedata.sizingdata[cell.title][`row_${cell.pos[1]}`] = sizedata
    }
    if (clearActSize) {
        this.setActSizingDefaultData()
    }
    this.updateproposalRowdata(storedata)
    // if (_.isEmpty(this.itemcodeexceldata)) {
    //   _.forEach(this.actuatorItemcodeList, (code) => {
    //     this.itemcodeexceldata.push({
    //       itemcode: code,
    //       Model: formdata.model
    //     })
    //   })
    // }
  }

  unmask(val) {
    // NOTE unmask
    if ((typeof val) !== 'string') {
      val = String(val)
    }

    val = val.replace(this.thousandSeparator || ',', '')
    val = val.replace(this.decimalSeparator, '.')
    val = parseFloat(val || 0)

    return val
  }

  /**
   * @method changed
   * @event onchange
   * @param { Object } instance
   * @param { Object } cell
   * @param { number } x
   * @param { number } y
   * @param { number|text|object|array } value
   * @param { number|text|object|array } oldvalue
   * @return {void}
   * @document  trigger on change of a cell datq
   */
  async changed(instance, cell, x, y, value, oldvalue) {
    console.log(this.rowpaste,'rowpaste')
      if(this.rowpaste) {
        console.log(x, value)
      const instancecolumns = instance.jexcel.options.columns
      const indexSysItemCode = _.findIndex(instancecolumns, 'isSystem')
      const sysColumnNameTarget = jexcel.getColumnNameFromId([indexSysItemCode, y])
      let multipleCode = ''
      // console.log(this.curColumn?.name, 'this.curColumn')
      const currentCol = instancecolumns?.[Number(x)]
      let indexQty = _.findIndex(this.jExcelOptions.columns, 'qty')
      if (indexQty === -1) {
        indexQty = _.findIndex(this.jExcelOptions.columns, { id: 'Qty' })
      }
      const qtyValue = parseInt(instance.jexcel.getValueFromCoords(indexQty, +y)) || 0

      if (value === 'add-new' && this.curColumn.productcol) {
        return
      }

      if (currentCol.price && value !== undefined && value !== null) {
        value = this.unmask(value)
      }

      if (instance.jexcel.options.columns[x].id === 'Product') {
        if (!qtyValue) {
          instance.jexcel.setValueFromCoords(indexQty, +y, 1)
        }
      }
      if (currentCol.multiple) {
        if (currentCol.multiple && +currentCol?.codeLen) {
              // const { id, cLength: codeLen } = currentCol || {}

              const derivedColumnId = +currentCol.derivedId
              const cLength = +currentCol.codeLen
              const dStart = (cLength === 3) ? 'NNN': 'NN'
              let derivedColumnValues = []
              let queryHeader = ''
              let query = ''
              let variables = {}
              queryHeader += `
                $cLength${currentCol.id}: Int!
                $dStart${currentCol.id}: String!
                $derivedColumnId${currentCol.id}: Int!
                $value${currentCol.id}: String!
                $depVal${currentCol.id}: String
              `
              query += `derivedCode${currentCol.id}: getDerivedCode(
                cLength: $cLength${currentCol.id}
                dStart: $dStart${currentCol.id}
                derivedColumnId: $derivedColumnId${currentCol.id}
                value: $value${currentCol.id}
                depVal: $depVal${currentCol.id}
              )`
              variables[`cLength${currentCol.id}`] = cLength
              variables[`dStart${currentCol.id}`] = dStart
              variables[`derivedColumnId${currentCol.id}`] = derivedColumnId
              // NOTE Step 3
              variables[`value${currentCol.id}`] = _.split(value,';').sort().join('')

              variables[`depVal${currentCol.id}`] = null
          if (!_.isEmpty(variables) && query) {
            const getDerivedCodeQuery = gql`query getDerivedCode(
              ${queryHeader}
            ) {
              ${query}
            }`

            await this.$apollo.query({
              variables,
              query: getDerivedCodeQuery
              // fetchPolicy: 'network-only'
            }).then(({ data }) => {
              // console.log(data)
              const result = _.keys(data)
              _.forEach(result, (key) => {
                const derivedCode = data[key]

                if (!_.isEmpty(derivedCode)) {
                  const { code } = derivedCode
                  multipleCode =  code
                  // instance.jexcel.setValue(sysColumnNameTarget, displayFgCode1)
                }
              })
            }).catch((error) => {
              this.$q.notify({
                type: 'info',
                message: error?.message || 'Something went wrong.!'
              })
            })
          }
        }
      }
      if (
        currentCol.multiple ||
        (value?.length ||
        Number(value) > 0 ||
        currentCol.price ||
        currentCol.discount || currentCol?.type === 'dropdown')
      ) {
        const columntemplate = _.cloneDeep(this.storedummydata?.columns || [])
        const productvalue = instance.jexcel.getValueFromCoords(
          _.findIndex(this.jExcelOptions.columns, {id: 'Product'}),
          y
        )
        const productTypeValue = instance.jexcel.getValueFromCoords(
          _.findIndex(this.jExcelOptions.columns, ['sourceId.name', 'Product Type']),
          y
        )
        const seriesValue = instance.jexcel.getValueFromCoords(
          _.findIndex(this.jExcelOptions.columns, ['sourceId.name', 'SERIES']),
          y
        )

        if (currentCol.id && currentCol.title) {
          const curColumnName = this.jExcelOptions.columns[Number(x)]?.id
          const fgtypelist = instance.jexcel.options?.meta?.fgCodeType
          const currfgtype = _.find(
            fgtypelist,
            fgrow => _.indexOf(fgrow.rows, y) > -1
          )

          if (
            currfgtype &&
            this.rowvalveprice &&
            currentCol.itemcode &&
            !_.isEmpty(itemcodelist) &&
            !_.isNull(this.rowvalveprice) &&
            value.split('-')?.length === currfgtype?.columns?.length
          ) {
            const valvelpindex = _.findIndex(instancecolumns, z => z.valveprice)

            if (valvelpindex !== -1) {
              const rowvalveprice = _.clone(this.rowvalveprice)
              this.$set(this, 'rowvalveprice', null)
              instance.jexcel.setValueFromCoords(
                valvelpindex,
                y,
                Number(rowvalveprice),
                true
              )
            }
          }

          if (
            currentCol.qty ||
            currentCol.price ||
            currentCol.unitval ||
            currentCol.testing ||
            currentCol.discount ||
            currentCol.finalval ||
            currentCol.valveprice
          ) {
            this.setfinalprice(instance, x, y, value)
          }

          if (value === 'NS:-NS' || value === 'Special Requirement:-S') {
            cell.style.cssText =
              'text-align: center; background-color: rgb(255, 245, 157);'
            nsOptions = []
            this.nonStandardList = []
            this.nsMdl = true
            this.nsColumnName = instance.jexcel.options.columns[x]
          } else {
            cell.style.cssText =
              'text-align: center; background-color: rgb(255, 255, 255);'
          }

          let rowmeta
          const fgTypeMeta = instance.jexcel.options?.meta?.fgCodeType

          if (fgTypeMeta) {
            rowmeta = _.find(fgTypeMeta, x => _.indexOf(x.rows, Number(y)) !== -1)
          }
          if (
            rowmeta &&
            currentCol.productcol &&
            oldvalue?.length &&
            value !== oldvalue &&
            !_.find(rowmeta?.product, ['id', value])
          ) {
            this.resetrow(value, y)
            return undefined
          }

          // NOTE listing all fgcode related columns
          let srcValue = null
          let arrofFGCells, fgvalue, fgtype
          const validprocuct = _.find(columntemplate?.fgCodeType, prod =>
            _.includes(_.map(prod.product, 'name'), productvalue)
          )

          if (!currentCol.itemcode && currentCol.isFgColumn) {
            if (
              validprocuct &&
              productvalue?.length &&
              columntemplate?.fgCodeType?.length
            ) {
              fgtype = validprocuct
              arrofFGCells = _.map(fgtype?.columns, 'id')
              fgvalue = [{ ...fgtype, rows: [Number(y)] }]
            } else if (columntemplate?.fgCodeType?.length) {
              if (instancecolumns[x].productDependent) {
                srcValue = _.find(instancecolumns[x].source, ['id', value])
                fgtype = srcValue
                  ? _.find(columntemplate?.fgCodeType, prod => {
                    return _.find(prod.product, ['id', srcValue?.Product?.[0]])
                  })
                  : columntemplate.fgCodeType[0]
              } else fgtype = columntemplate.fgCodeType[0]
              arrofFGCells = _.map(fgtype?.columns, 'id')
              fgvalue = [{ ...fgtype, rows: [Number(y)] }]
            }

            const derievedcolList = _.flatMap(
              columntemplate.derivedColumn,
              'columns'
            )
            // let arrofFGCells1 = _.cloneDeep(arrofFGCells).map(x=>x.split('_').join(''))
            if (
              arrofFGCells &&
              (_.includes(arrofFGCells, curColumnName) || _.includes(_.map(derievedcolList, 'id'), curColumnName))
            ) {
              if (!_.isEmpty(instance.jexcel.options?.meta?.['fgCodeType'])) {
                const fgtypes = _.isArray(
                  instance.jexcel.options.meta['fgCodeType']
                )
                  ? instance.jexcel.options.meta['fgCodeType']
                  : _.values(instance.jexcel.options.meta['fgCodeType'])
                const currFgTypeIndex = _.findIndex(fgtypes, ['id', fgtype?.id])
                const currentfgType = fgtypes[currFgTypeIndex]
                const oldfgTypeIndex = _.findIndex(fgtypes, x =>
                  _.includes(x.rows, Number(y))
                )

                if (oldfgTypeIndex !== -1) {
                  fgtypes[oldfgTypeIndex].rows.splice(
                    _.indexOf(fgtypes[oldfgTypeIndex]?.rows, Number(y)),
                    1
                  )
                }

                if (currentfgType) {
                  currentfgType.rows = [...currentfgType.rows, Number(y)]
                  fgtypes[currFgTypeIndex] = currentfgType
                } else {
                  fgtypes.push({ ...fgtype, rows: [Number(y)] })
                }
                fgvalue = fgtypes
              }
              instance.jexcel.options.meta['fgCodeType'] = fgvalue
              this.itemcodereation(
                instance,
                cell,
                x,
                y,
                value,
                currentCol,
                arrofFGCells,
                multipleCode
              )
              const productval = instance.jexcel.getValueFromCoords(
                _.findIndex(instancecolumns, 'productcol'),
                y
              )

              if (
                !productval?.length &&
                currentCol.productDependent
              ) {
                instance.jexcel.setValueFromCoords(
                  _.findIndex(instancecolumns, 'productcol'),
                  y,
                  srcValue?.Product?.[0]
                )
              }
            }
          }
      if(this.valepopsetting && (this.company?.code === '71')) {
          console.log(this.valepopsetting, 'valepopsetting')
          let valveName = 'Valve P/N'  // VALVE P/N
          if (this.company?.code === '71') {
            valveName = 'VALVE P/N'
          }
          const valveMeta =  instance.jexcel.options.meta[valveName]
          if ((valveMeta !== undefined ) || ( instance.jexcel.options.meta?.[valveName]?.[y] !== undefined)) {
          const fgcolmeta = _.findIndex(instance.jexcel.options.meta[valveName]?.[y]?.productMeta?.Fgcolumns, { id: currentCol.id })
          if( fgcolmeta > -1) {
            console.log(currentCol.id, 'currCol')
          let newFgValue  =  _.find(_.find(instance.jexcel.options?.meta[valveName]?.[y]?.productMeta?.Fgcolumns, { id: currentCol.id })?.source, { id: value } )
          let nwCode = ''
          let newval = ''
          if (currentCol.multiple){
            newFgValue = value
            nwCode = multipleCode
            newval = value
          } else {
            nwCode = _.includes(value, ':-') ? value?.split(':-')[1] : ''
            newval = _.includes(value, ':-') ? value?.split(':-')[0] : value
          }
          console.log(newFgValue, 'newFg')
          this.$set(instance.jexcel.options.meta[valveName]?.[y]?.productMeta.Fgcolumns[fgcolmeta], 'fgValue' , newFgValue)
          valveMeta[y].productMeta.itemcodearray[fgcolmeta] = nwCode
          valveMeta[y].productMeta.shortDesArr[fgcolmeta] = newval
          valveMeta[y].productMeta.longDesArr[fgcolmeta] = newval
          valveMeta[y].productMeta.longDes =  instance.jexcel.options.meta[valveName][y].productMeta.longDesArr.join(', ')
          valveMeta[y].productMeta.shortDes =  instance.jexcel.options.meta[valveName][y].productMeta.shortDesArr.join(', ')
          valveMeta[y].productMeta.FGCode = instance.jexcel.options.meta[valveName][y].productMeta.itemcodearray.join('')
          valveMeta[y].productMeta.desc = instance.jexcel.options.meta[valveName][y].productMeta.shortDes
          valveMeta[y].productMeta.descLg = instance.jexcel.options.meta[valveName][y].productMeta.longDes
          valveMeta[y].productMeta.itemcode =  instance.jexcel.options.meta[valveName][y].productMeta.FGCode
          valveMeta[y].desc = instance.jexcel.options.meta[valveName][y].productMeta.shortDes
          valveMeta[y].descLg = instance.jexcel.options.meta[valveName][y].productMeta.longDes
          valveMeta[y].itemcode =  instance.jexcel.options.meta[valveName][y].productMeta.FGCode
      }
    }
  }
          // console.log(currentCol, 'currentCol.targetColumn')
          if (currentCol.targetColumn && currentCol.multiplier) {
            this.setTarget(instance, y, value, currentCol)
          }
        }

        if (
          !currentCol.multiple &&
          value &&
          value !== '-' &&
          value !== '--:-' &&
          seriesValue !== '-' &&
          seriesValue !== 'OS:-OS' &&
          !_.includes(value, ':-NS') &&
          !_.includes(seriesValue, ':-NS') &&
          !_.includes(seriesValue, '--:-') &&
          value !== 'Special Requirement:-S' &&
          instance.jexcel.options.columns[x].type === 'dropdown' &&
          !_.includes(['product_type', 'product', 'product group'], _.toLower(currentCol.title))
        ) {
          let source = instance.jexcel.options.columns[x].source
          const [dependency] = instance.jexcel.options.columns[x]?.dependenciesId || []

          if (_.includes(_.toLower(dependency?.name), 'series') && seriesValue) {
            const [_seriesLbl, series] = _.split(seriesValue, ':-')
            source = _.filter(source, x => _.includes(x.Series, series))
          } else if (_.includes(_.toLower(dependency?.name), 'product type') && productTypeValue) {
            source = _.filter(source, x => _.includes(x['Product type'], productTypeValue))
          } else if (_.includes(_.toLower(instance.jexcel.options.columns[x].id), 'series') && productvalue) {
            source = _.filter(source, x => _.includes(x.Product, productvalue))
          }

          if (!_.find(source, { id: value })) {
            instance.jexcel.setValueFromCoords(
              _.findIndex(instance.jexcel.options.columns, { title: instance.jexcel.options.columns[x].title }),
              y,
              null
            )
          }
        }
      }
    }
  }

  /**
   * @method setTarget
   * @param {String} product
   * @return {void}
   * @document reset Row data
   */
  async setTarget(instance, y, value, currentCol) {
    let product = Number(value) || 0
    _.forEach(currentCol.multiplier, m => {
      const cellval =
        Number(
          instance.jexcel.getValueFromCoords(
            _.findIndex(instance.jexcel.options.columns, m),
            y
          )
        ) || 0
      product = cellval * product
    })

    if (product) {
      await instance.jexcel.setValueFromCoords(
        _.findIndex(instance.jexcel.options.columns, currentCol.targetColumn),
        y,
        product
      )
    }
  }

  /**
   * @method setfinalprice
   * @param {String} product
   * @return {void}
   * @document reset Row data
   */
  setfinalprice(instance, x, y, value) {
    let unitcost = 0
    const pricecol = []
    const col = Number(x)
    let finalunitcost = 0
    const valofy = Number(y)
    const priceindexlist = []
    const metaData = instance.jexcel.getMeta()
    const currentCol = instance.jexcel.options.columns[Number(x)]
    const indexQty = _.findIndex(this.jExcelOptions.columns, 'qty')
    const indexunit = _.findIndex(this.jExcelOptions.columns, 'unitval')
    const indexfinal = _.findIndex(this.jExcelOptions.columns, 'finalval')
    const indextotal = _.findIndex(this.jExcelOptions.columns, 'totalval')
    const productColumn = _.findIndex(this.jExcelOptions.columns, 'productcol')
    const qtyvalue = parseInt(instance.jexcel.getValueFromCoords(indexQty, valofy)) || 0

    if (_.includes(value, this.currency)) {
      value = _.split(value, this.currency)[1]
    }

    if (
      currentCol.price &&
      !currentCol.unitval &&
      !currentCol.totalval &&
      !currentCol.testing &&
      !currentCol.finalval
    ) {
      _.forEach(this.jExcelOptions.columns, (c, i) => {
        if (c.price && !c.unitval && !c.totalval && !c.testing && !c.finalval) {
          priceindexlist.push(i)
        }
      })

      const valuelist = []
      _.forEach(priceindexlist, x => {
        pricecol.push(this.jExcelOptions.columns[x].id)

        let a = instance.jexcel.getValueFromCoords(x, y)
        a = _.includes(a, this.currency) ? _.split(a, this.currency)[1] : a

        if (a !== undefined && a !== null) {
          a = this.unmask(a)
        }

        a = isNaN(Number(a)) ? 0 : Number(a)
        a = x === col ? (isNaN(Number(value)) ? 0 : Number(value)) : a

        if (a > 0) {
          valuelist.push(a)
        }
      })

      unitcost = _.round(_.sum(valuelist), 2)
    }

    if (
      currentCol.unitval ||
      currentCol.discount
    ) {
      const discount = _.findIndex(this.jExcelOptions.columns, 'discount')
      let discountval = parseFloat(
        instance.jexcel.getValueFromCoords(discount, valofy)
      )
      unitcost = currentCol.unitval
        ? Number(value)
        : unitcost === 0
          ? parseFloat(instance.jexcel.getValueFromCoords(indexunit, valofy))
          : unitcost

      if (!discountval) {
        discountval = 0
      }

      let outSourcedprice = 0
      _.map(this.jExcelOptions.columns, (c, index) => {
        if (
          c.price &&
          !c.unitval &&
          !c.totalval &&
          !c.testing &&
          !c.finalval &&
          metaData[c.title]?.[y]?.outSource
        ) {
          let a = instance.jexcel.getValueFromCoords(index, y)

          if (a !== undefined && a !== null) {
            a = this.unmask(a)
          }

          a = _.includes(a, this.currency) ? _.split(a, this.currency)[1] : a
          a = isNaN(Number(a)) ? 0 : Number(a)
          a = x === index ? (isNaN(Number(value)) ? 0 : Number(value)) : a

          if (a > 0) outSourcedprice += a
        }
      })

      if (this.company.code === '03_DELVAL') {
        finalunitcost = discountval
          ? (unitcost - ((unitcost - outSourcedprice) * (discountval / 100)))
          : unitcost
      } else if (this.company.code === '71') {
        finalunitcost = discountval
          ? (unitcost - ((unitcost - outSourcedprice) * (discountval / 100)))
          : unitcost
      }
      else {
        const miltiplier = discountval || 1
        finalunitcost = (outSourcedprice + ((unitcost - outSourcedprice) * miltiplier))
      }

      instance.jexcel.setValueFromCoords(
        indexfinal,
        valofy,
        parseFloat(finalunitcost || 0).toFixed(this.roundCount),
        true
      )
    } else if (_.includes(priceindexlist, col)) {
      instance.jexcel.setValueFromCoords(
        indexunit,
        valofy,
        unitcost.toFixed(this.roundCount),
        true
      )
    } else if (currentCol.qty || currentCol.finalval || currentCol.testing) {
      const indexTesting = _.findIndex(this.jExcelOptions.columns, 'testing')
      const finalrate = currentCol.finalval
        ? Number(value)
        : parseFloat(instance.jexcel.getValueFromCoords(indexfinal, valofy)) ||
          0
      let testingcharges = currentCol.testing
        ? Number(value)
        : parseFloat(
          instance.jexcel.getValueFromCoords(indexTesting, valofy)
        ) || 0
      const totalcost = _.round(finalrate * qtyvalue, 2) + _.round(testingcharges * qtyvalue, 2)
      testingcharges = testingcharges || 0
      instance.jexcel.setValueFromCoords(
        indextotal,
        valofy,
        parseFloat(totalcost, 1).toFixed(this.roundCount),
        true
      )
    }
    let jsonData = _.filter(instance.jexcel.getJson(), 'Product')
    if (
      !jsonData?.length ||
      !instance.jexcel.getValueFromCoords(productColumn, y)
    ) {
      jsonData = instance.jexcel.getJson()
      this.showWarningMsg({ message: 'Please select Product for Current Row' })
    }
    if (indextotal !== -1) {
      const totalval1 = _.map(
        jsonData,
        tval =>
          parseFloat(tval[this.jExcelOptions.columns[indextotal].title]) || 0
      )

      const emitobj = {
        qty: _.sum(
          _.map(
            jsonData,
            b =>
              parseFloat(b[instance.jexcel.options.columns[indexQty].title]) ||
              0
          )
        ),
        total: _.sum(totalval1)
      }

      this.$emit('sumdata', emitobj)
    }
  }

  itemcodereation(instance, cell, x, y, value, currentCol, arrofFGCells, multipleCode) {
    let oldTrimVal = []
    const sheetoptions = instance.jexcel.options
    const instancecolumns = sheetoptions.columns
    const curColumnName = currentCol?.id
    const columntemplate = this.storedummydata.columns
    var fgcodePositionIndexArr = _.map(arrofFGCells, (z, i) => i)
    let targetCell = ''
    const indexItemCode = _.findIndex(instancecolumns, 'itemcode')

    var columnNameTarget = jexcel.getColumnNameFromId([indexItemCode, y])

    if (columnNameTarget) {
      targetCell = instance.jexcel.getCell(columnNameTarget)
    }
    // if (sysColumnNameTarget) {
    //   systargetCell = instance.jexcel.getCell(sysColumnNameTarget)
    // }
    // console.log(indexSysItemCode, sysColumnNameTarget)
    var oldFgCode = ['', '', '', '', '', '', '']
    const fgCodeType = _.find(
      sheetoptions?.meta?.fgCodeType,
      c => _.indexOf(c.rows, Number(y)) > -1
    )

    let indextoFgcode = _.indexOf(arrofFGCells, curColumnName)

    if (columntemplate?.derivedColumn?.length) {
      let fgTrimcodeColIndex = -1
      const trimCodeCols = _.filter(
        columntemplate.derivedColumn,
        (z, index) => {
          return _.find(z.columns, ['id', curColumnName])
        }
      )
      const trimCodeCol = _.find(trimCodeCols, x =>
        _.find(fgCodeType?.columns, ['title', x.title])
      )
      if (trimCodeCol) {
        indextoFgcode = _.indexOf(
          arrofFGCells,
          _.replace(trimCodeCol.title, / /g, '_')
        )
        if (indextoFgcode === -1) {
          indextoFgcode = _.indexOf(arrofFGCells, trimCodeCol.title)
        }
        fgTrimcodeColIndex = _.findIndex(trimCodeCol.columns, [
          'id',
          curColumnName
        ])
        const trimMeta = instance.jexcel.getMeta()

        if (
          trimMeta.ItemCode &&
          trimMeta.ItemCode[`${trimCodeCol?.title}_${y}`]
        ) {
          let ObjArr = {}

          ObjArr = trimMeta.ItemCode[`${trimCodeCol?.title}_${y}`]
          oldTrimVal = ObjArr === undefined ? [] : ObjArr
        }
        oldTrimVal[fgTrimcodeColIndex] = _.split(value || '', ':-')?.[0]

        instance.jexcel.setMeta(
          'ItemCode',
          `${trimCodeCol?.title}_${y}`,
          oldTrimVal
        )
        var fgtcode = null
        var fgtcodeObj = {}
        if (
          _.filter(oldTrimVal, t => !_.isEmpty(t))?.length ===
          trimCodeCol?.columns?.length
        ) {
          fgtcode = '000'
          const tobj = {}
          _.forEach(trimCodeCol.columns, (c, i) => {
            tobj[_.toLower(c.sourceId.name)] = oldTrimVal[i]
          })
          fgtcodeObj = _.find(trimCodeCol?.source, tobj)

          fgtcodeObj = fgtcodeObj === undefined ? {} : fgtcodeObj
          fgtcode = fgtcodeObj?.fgtCode || fgtcode

          const Obbj = instance.jexcel.getMeta()

          if (Obbj.ItemCode) {
            const ObjArr = Obbj.ItemCode

            // get old  FGcode and update

            oldFgCode =
              ObjArr[`fgcode_${y}`] === undefined ? [] : ObjArr[`fgcode_${y}`]
          }
          const derivedindex = _.indexOf(
            arrofFGCells,
            _.replace(trimCodeCol?.title, / /g, '_')
          )

          if (fgtcode === '000') {
            this.showInfoMsg({ message: 'Trim Code not available' })
          }
          oldFgCode[derivedindex] = fgtcode
          instance.jexcel.setMeta('ItemCode', `fgcode_${y}`, oldFgCode)

          var displayFgCode1 = oldFgCode.join('-')
          if (targetCell) {
            targetCell?.classList.remove('readonly')
            // systargetCell?.classList.remove('readonly')
            // console.log(displayFgCode1, 'displayFgCode1')

            instance.jexcel.setValue(columnNameTarget, displayFgCode1)
            // instance.jexcel.setValue(sysColumnNameTarget, displayFgCode1)

            targetCell.classList.add('readonly')
            // systargetCell.classList.add('readonly')
          }
        }
      }
    }
    const fgcodePositionIndex = fgcodePositionIndexArr[indextoFgcode]

    if (value !== 'NS:-NS' || !currentCol.multiple) {
      cell.innerText = _.split(value || '', ':-')?.[0]
    }

    const Obbj = instance.jexcel.getMeta()
    oldFgCode = ['', '']
    if (
      Obbj.ItemCode &&
      _.includes(Object.keys(Obbj.ItemCode), `fgcode_${y}`)
    ) {
      let ObjArr = {}
      ObjArr = Obbj.ItemCode

      oldFgCode = ObjArr[`fgcode_${y}`] || ['', '']
    }
    if (arrofFGCells[fgcodePositionIndex] === curColumnName) {
      if(currentCol.multiple){
        oldFgCode[fgcodePositionIndex] = multipleCode?.trim()
      } else {
      oldFgCode[fgcodePositionIndex] = _.split(value || '', ':-')?.[1]
      }
    }
    if (oldFgCode.length === arrofFGCells.length) {
      const codelist = _.reduce(
        fgCodeType?.columns,
        (list, c, i) => {
          if (c?.coltype?.id === 1) list.push(oldFgCode[i])
          return list
        },
        []
      )

      this.rowvalveprice = itemcodelist[_.join(codelist, '-')]
        ? itemcodelist[_.join(codelist, '-')]
        : 0
    }

    instance.jexcel.setMeta('ItemCode', `fgcode_${y}`, oldFgCode)
    const oldObjValveLp = instance.jexcel.getMeta()?.Valve_LP?.[y]
    if(!_.isUndefined(oldObjValveLp)) {
    const qtyIndex =  _.findIndex(this.jExcelOptions.columns, {id: 'Qty.'})
    const qtyval = instance.jexcel.getValueFromCoords(qtyIndex, y)
    const vlvLpIndex = _.findIndex(this.jExcelOptions.columns, {id: 'valve_price'})
    const vlvPnIndex = _.findIndex(this.jExcelOptions.columns, {id: 'Valve_P/N'})
    const oldSDesc = instance.jexcel.getMeta().Valve_LP[y]?.desc
    let oldDescArr = oldSDesc.split(',')
    if (_.isArray(oldDescArr)) {
    const changedAttr = oldDescArr[fgcodePositionIndex].split(':-')[0]
    const newAttrDesc = `${changedAttr}:-${value}`
    oldDescArr[fgcodePositionIndex] = newAttrDesc
    const newDescArr = oldDescArr
    oldObjValveLp['desc'] = _.isArray(newDescArr) ? newDescArr.join(',') : newDescArr
    instance.jexcel.setComments([vlvLpIndex, y],  oldObjValveLp.desc)
    instance.jexcel.setComments([vlvPnIndex, y],  oldObjValveLp.desc)
    }
    oldObjValveLp['price'] = this.rowvalveprice
    oldObjValveLp['qty'] = qtyval

    instance.jexcel.setMeta('Valve P/N', y, oldObjValveLp)
    instance.jexcel.setMeta('Valve_LP', y, oldObjValveLp)
  }

    // update the meta valve_PN
    // get meta -> get current value and update with new value


    const displayFgCode = (
      _.includes(oldFgCode, 'NS') ||
      _.includes(oldFgCode, 'OS') ||
      _.includes(oldFgCode, 'NA')
    ) ? (this.company?.code === '03_DELVAL' ? 'CTBG' : 'TBD')
      : _.join(oldFgCode, '-')
    targetCell = instance.jexcel.getCell(columnNameTarget)

    if (targetCell) {
      targetCell.classList.remove('readonly')
      // systargetCell.classList.remove('readonly')

      instance.jexcel.setValue(columnNameTarget, displayFgCode)
      // instance.jexcel.setValue(sysColumnNameTarget, displayFgCode)
      targetCell.classList.add('readonly')
      // systargetCell.classList.add('readonly')
    }
  }

  async syscodeCreation () {
    const sheet = this.$refs.spreadsheet.jexcel
    const instance = sheet.el
    const sheetoptions = instance.jexcel.options
    const instancecolumns = sheetoptions.columns
    const y = this.currentcell?.pos?.[1]
    if (this.company?.code === '71') {
    const sysCol = _.filter(instancecolumns, { isSysColumn: true })
    const valvCol = _.orderBy(_.filter(sysCol, x =>  { if(!x.productId?.id) return x} ), 'index')
    const valveMetaCol = this.currentcell?.instance?.jexcel?.options?.meta['VALVE P/N']?.[y]?.productMeta?.Fgcolumns
    const ValvCodeArr = this.currentcell?.instance?.jexcel?.options?.meta?.ItemCode?.fgcode_0
    const prodCol = _.orderBy(_.filter(sysCol, x =>  { if(x.productId?.id > 0) return x} ), 'index')
    // seperate valve columns
    // from meta get Code
    // product columns, find sys attr and get code
    let sysPrCode = ['VC','-','0', '00000000000', '-', '0', '000', '-', '00','-', '00', '-', '00', '-', '00', '0', '0','-', '0']
    let sysPrShortDes = []
    let sysUnqieValDerCode = []
    let produtcTypeColIndex= _.filter(sysCol, {title:'Product Type'})?.[0]?.index
    let produTypeVal = instance.jexcel.getValueFromCoords(produtcTypeColIndex, y)
    const productTypeCode =  produTypeVal?.split(':-')?.[1]
    // if (_.includes(['x'], productTypeCode)) {
    //   sysPrCode = ''
    //   sysPrShortDes = ''
    //   let systargetCell = ''
    //   const indexSysItemCode = _.findIndex(instancecolumns, 'isSystem')
    //   const sysColumnNameTarget = jexcel.getColumnNameFromId([indexSysItemCode, y])
    //   if (sysColumnNameTarget) {
    //     const cell = this.currentcell
    //     debugger
    //     systargetCell = instance.jexcel.getCell(sysColumnNameTarget)
    //     systargetCell?.classList.remove('readonly')
    //     instance.jexcel.setValue(sysColumnNameTarget, sysPrCode)
    //     systargetCell?.classList.add('readonly')
    //     cell.instance.jexcel.options.meta = {
    //       ...cell.instance.jexcel.options.meta,
    //       ['SysItemCode']: {
    //         ...cell.instance.jexcel.options.meta['SysItemCode'],
    //         [cell.pos[1]]: { ['itemcode'] : sysPrCode, ['descLg']: shortsysDesc }
    //       }
    //     }
    //   }
    // }
      sysPrCode[2] = (produTypeVal?.split(':-')?.[1])
      sysUnqieValDerCode.push(produTypeVal?.split(':-')?.[1])
      let protypeVal = produTypeVal?.split(':-')?.[0]
      sysPrShortDes.push(`${protypeVal?.toUpperCase()}, \n`)
      // val 11 digit
      let sysValPrCode = []
     let valCols = _.intersectionBy(valveMetaCol, valvCol, 'index')
     sysUnqieValDerCode.push(this.currentcell?.instance?.jexcel?.options?.meta['VALVE P/N']?.[y]?.itemcode)
    let fgInstCols = _.filter(instancecolumns, col => { if (col?.productId && col.productId?.id > 0) return col})
     _.forEach(fgInstCols, (prcol)=> {
     sysUnqieValDerCode.push(this.currentcell?.instance?.jexcel?.options?.meta[prcol.title]?.[y]?.itemcode)
     })
     _.forEach(valCols, x => {
      // console.log(x.title)
      let sysValCode = x?.fgValue?.name?.split(':-')[1]
      if (_.includes(sysValCode,'VC')){
        sysValCode = sysValCode.split('-')[1]
      }
      sysValPrCode.push(sysValCode)
      if(!_.includes(['VALVE SERIES', 'BRAND', 'VALVE SIZE', 'BORE', 'RATING'], x?.title)) {
        sysPrShortDes.push(`${x?.title}:-`)
        sysPrShortDes.push(`${x?.fgValue?.name?.split(':-')[0]}, `)
        } else {
        sysPrShortDes.push(`${x?.fgValue?.name?.split(':-')[0]}, `)
        }
    })
    sysPrCode[3] = sysValPrCode.join('')
    _.forEach(prodCol, (prcol)=> {
     let valCols = _.filter(this.currentcell?.instance?.jexcel?.options?.meta[prcol.title]?.[y]?.productMeta?.Fgcolumns, { isSysColumn: true })
        // console.log(valCols.length, prcol.title, prcol)
        let colcode = []
        if (valCols.length) {
      _.forEach(valCols, x => {
        let vscolval = x?.fgValue?.name?.split(':-')[1]
        if (_.includes(vscolval,'VC')){
          vscolval = `${vscolval.split('-')[1]}`
      }
        colcode.push(vscolval)
        if(!_.includes(['ACTUATOR SERIES'], x?.title)) {
          sysPrShortDes.push(` ${x?.title}:-`)
          sysPrShortDes.push(` ${x?.fgValue?.name?.split(':-')[0]},`)
        } else {
          sysPrShortDes.push(` ${x?.fgValue?.name?.split(':-')[0]}`)
       }
       console.log(valCols)
      })
      if (_.includes(['HAND LEVER', 'GR.Box/MOR'], prcol.title)) {
        sysPrCode[5]= colcode?.join('')
      } else  if (_.includes(['PNEUMATIC ACTUATOR', 'MOTORIZED ACTUATOR'], prcol.title)) {
        sysPrCode[6]= colcode?.join('')
      } else  if (_.includes(['MAIN SOV'], prcol.title)) {
        sysPrCode[8]= colcode?.join('')
      } else  if (_.includes(['FCV', 'QEV'], prcol.title)) {
        sysPrCode[10]= colcode?.join('')
      }  else  if (_.includes(['POSITION INDICATOR', 'POSITIONER'], prcol.title)) {
        sysPrCode[12]= colcode?.join('')
      } else if (prcol.title === 'AFR') {
        sysPrCode[14] = colcode?.join('')
      } else if (prcol.title === 'CABLE GLANDS & CABLING') {
        sysPrCode[15] = colcode?.join('')
      } else if (prcol.title === 'TUBING') {
        sysPrCode[16]= colcode?.join('')
      }
    } else {
    }
    })

    sysPrCode = sysPrCode.join('')
    // console.log(sysPrShortDes.join(''))
    let shortsysDesc = sysPrShortDes.join('')
    shortsysDesc = shortsysDesc.replace(/(^,)|(,$)/g, "")
     sysUnqieValDerCode = _.compact(sysUnqieValDerCode)
    let sysVal = sysUnqieValDerCode.join('')
    // console.log(sysVal)
    if (sysUnqieValDerCode.length > 1) {
      if (sysUnqieValDerCode.length > 1) {
            // const { id, cLength: codeLen } = currentCol || {}
            const sysId = 4
            const derivedColumnId = 4
            const cLength = 3
            const dStart = 'NNN'
            let derivedColumnValues = []
            let queryHeader = ''
            let query = ''
            let variables = {}
            queryHeader += `
              $cLength${sysId}: Int!
              $dStart${sysId}: String!
              $derivedColumnId${sysId}: Int!
              $value${sysId}: String!
              $depVal${sysId}: String
            `
            query += `derivedCode${sysId}: getDerivedCode(
              cLength: $cLength${sysId}
              dStart: $dStart${sysId}
              derivedColumnId: $derivedColumnId${sysId}
              value: $value${sysId}
              depVal: $depVal${sysId}
            )`
            variables[`cLength${sysId}`] = cLength
            variables[`dStart${sysId}`] = dStart
            variables[`derivedColumnId${sysId}`] = derivedColumnId
            // NOTE Step 3
            variables[`value${sysId}`] = sysVal
            variables[`depVal${sysId}`] = sysPrCode

        if (!_.isEmpty(variables) && query) {
          const getDerivedCodeQuery = gql`query getDerivedCode(
            ${queryHeader}
          ) {
            ${query}
          }`

          await this.$apollo.query({
            variables,
            query: getDerivedCodeQuery
            // fetchPolicy: 'network-only'
          }).then(({ data }) => {
            // console.log(data)
            const result = _.keys(data)
            _.forEach(result, (key) => {
              const derivedCode = data[key]

              if (!_.isEmpty(derivedCode)) {
                const { code } = derivedCode
                sysPrCode = sysPrCode + code
                let systargetCell = ''
                const indexSysItemCode = _.findIndex(instancecolumns, 'isSystem')
                const sysColumnNameTarget = jexcel.getColumnNameFromId([indexSysItemCode, y])
                if (sysColumnNameTarget) {
                  const cell = this.currentcell
                  systargetCell = instance.jexcel.getCell(sysColumnNameTarget)
                  systargetCell?.classList.remove('readonly')
                  instance.jexcel.setValue(sysColumnNameTarget, sysPrCode)
                  systargetCell?.classList.add('readonly')
                  cell.instance.jexcel.options.meta = {
                    ...cell.instance.jexcel.options.meta,
                    ['SysItemCode']: {
                      ...cell.instance.jexcel.options.meta['SysItemCode'],
                      [cell.pos[1]]: { ['itemcode'] : sysPrCode, ['descLg']: shortsysDesc }
                    }
                  }
                }
              }
            })
          }).catch((error) => {
            this.$q.notify({
              type: 'info',
              message: error?.message || 'Something went wrong.!'
            })
          })
        }
      }
    }
    //
  }
}

  codecreation(
    {
      sheet1,
      columntemplate,
      qtycol,
      finalval,
      lumpsum,
      totalval,
      discount,
      unitval1,
      discountval,
      productcol
    },
    row,
    index,
    comment
  ) {
    if (discount) {
      const valuelist = []

      _.forEach(this.jExcelOptions.columns, (c, i) => {
        let obj
        if (c.desc && c.desc === true) {
          if (comment[c.title]) {
            obj = { desc: comment[c.title], price: row[c.title] }
          } else if (row[c.title]) {
            obj = { desc: '', price: row[c.title] }
          }

          sheet1.setMeta(c.title, String(index - 1), obj)
        }

        if (c.price && !c.unitval && !c.totalval && !c.testing && !c.finalval) {
          const a = row[c.title]

          if (!isNaN(Number(a))) {
            valuelist.push(Number(a))
          }
        }
      })

      if (valuelist.length) {
        this.$set(row, unitval1, parseFloat(_.sum(valuelist) || 0).toFixed(this.roundCount))
      }

      // row[finalval] = row[unitval1] - (row[unitval1] * row[discountval]) / 100
      row[finalval] = row[unitval1] * row[discountval]
      if (this.company.code === '71') {
        row[finalval] = row[unitval1] - (row[unitval1] * (row[discountval]/100))
      }
    }

    row[totalval] = parseFloat(
      row?.[qtycol.title] * row[finalval] +
        (!row[lumpsum?.title] ? 0 : row[lumpsum.title]) * row?.[qtycol.title],
      2
    ).toFixed(this.roundCount)

    var arrofFGCells = []
    const productvalue = row[productcol]
    const validprocuct = _.find(columntemplate?.fgCodeType, prod =>
      _.includes(_.map(prod.product, 'id'), productvalue)
    )

    if (
      columntemplate?.fgCodeType?.length &&
      productvalue?.length &&
      validprocuct
    ) {
      const fgtype = validprocuct
      arrofFGCells = _.map(fgtype?.columns, 'id')
      let fgvalue = [{ ...fgtype, rows: [index - 1] }]
      if (sheet1.options.meta?.fgCodeType?.length) {
        const fgtypes = sheet1.options.meta['fgCodeType']
        const currFgTypeIndex = _.findIndex(fgtypes, ['id', fgtype?.id])
        const currentfgType = fgtypes[currFgTypeIndex]
        const oldfgTypeIndex = _.findIndex(fgtypes, x =>
          _.includes(x.rows, index - 1)
        )
        if (oldfgTypeIndex !== -1) {
          fgtypes[oldfgTypeIndex].rows.splice(
            _.findIndex(fgtypes[oldfgTypeIndex]?.rows, index - 1),
            1
          )
        }
        if (currentfgType) {
          currentfgType.rows = [...currentfgType.rows, index - 1]
          fgtypes[currFgTypeIndex] = currentfgType
        } else fgtypes.push({ ...fgtype, rows: [index - 1] })
        fgvalue = fgtypes
      }
      sheet1.options.meta['fgCodeType'] = fgvalue
    } else if (columntemplate?.fgCodeType?.length) {
      arrofFGCells = _.map(columntemplate.fgCodeType[0]?.columns, 'id')
      let fgvalue = [{ ...columntemplate.fgCodeType[0], rows: [index - 1] }]
      if (sheet1.options.meta['fgCodeType']) {
        const fgtypes = sheet1.options.meta['fgCodeType']
        const currFgTypeIndex = _.findIndex(fgtypes, [
          'id',
          columntemplate.fgCodeType[0]?.id
        ])
        const currentfgType = fgtypes[currFgTypeIndex]
        const oldfgTypeIndex = _.findIndex(fgtypes, x =>
          _.includes(x.rows, index - 1)
        )
        if (oldfgTypeIndex !== -1) {
          fgtypes[oldfgTypeIndex].rows.splice(
            _.findIndex(fgtypes[oldfgTypeIndex]?.rows, index - 1),
            1
          )
        }
        if (currentfgType) {
          currentfgType.rows = [...currentfgType.rows, index - 1]
          fgtypes[currFgTypeIndex] = currentfgType
        } else {
          fgtypes.push({ ...columntemplate.fgCodeType[0], rows: [index - 1] })
        }
        fgvalue = fgtypes
      }
      sheet1.options.meta['fgCodeType'] = fgvalue
    } else {
      arrofFGCells = _.map(columntemplate?.fgCode, 'columns.title')
      // this.jExcelObj.setMeta('fgCodeType', index - 1, undefined)
    }
    const fgval = []
    arrofFGCells.forEach(v => {
      if (_.includes(row[v], ':-')) {
        fgval.push(_.split(row[v], ':-')?.[1])
      } else fgval.push('')
    })
    const oldFgCode = fgval
    let fgtcodeObj = {}

    if (!_.isEmpty(columntemplate.derivedColumn)) {
      _.forEach(columntemplate.derivedColumn, col => {
        const tobj = {}
        const trimarray = []

        _.forEach(col.columns, (c, i) => {
          if (row[c.title]) {
            tobj[_.toLower(c.sourceId.name)] = _.split(row[c.title], ':-')?.[0]
            trimarray.push(_.split(row[c.title], ':-')?.[0])
          } else if (row[col] === '') {
            tobj[_.toLower(c.sourceId.name)] = row[c.title]
            trimarray.push(row[col])
          }
        })
        fgtcodeObj = _.find(col.source, tobj)

        if (fgtcodeObj === undefined) {
          fgtcodeObj = {}
        }

        if (_.includes(arrofFGCells, col.title)) {
          oldFgCode[_.indexOf(arrofFGCells, col.title)] = !fgtcodeObj?.fgtCode
            ? '000'
            : fgtcodeObj.fgtCode
          this.jExcelObj.setMeta(
            'ItemCode',
            `${col.title}_${Number(index) - 1}`,
            trimarray
          )
        }
      })
    }
    this.jExcelObj.setMeta('ItemCode', `fgcode_${Number(index) - 1}`, oldFgCode)
    const displayFgCode = oldFgCode.join('-')
    row['Itemcode'] = displayFgCode

    const rowaray = []
    const columntitles = {}

    _.forEach(this.jExcelObj?.options?.columns, x => {
      columntitles[x.title] = x.title

      if (x.itemcode) {
        rowaray.push(displayFgCode)
      } else rowaray.push(row[x.title] || '')
    })

    return rowaray
  }

  /**
   * @method resetrow
   * @param {String} product
   * @return {void}
   * @document reset Row data
   */
  resetrow(product, i) {
    const sheet = this.$refs.spreadsheet?.jexcel
    const rowdata = _.map(this.jExcelOptions.columns, x => {
      //
      // eslint-disable-next-line
      return x?.sourceId?.id == 90 ? product : ''
    })
    const metadata = sheet.options.meta
    const proposaldata = this.getProposaldata

    _.forEach(this.jExcelOptions.columns, z => {
      if (z.sizing) {
        if (
          proposaldata?.sizingdata?.[z.title] &&
          proposaldata.sizingdata[z.title][i]
        ) {
          delete proposaldata.sizingdata[z.title][i]
        }
      }
      if (metadata[z.title]) {
        if (z.title === 'ItemCode') {
          if (metadata[z.title]['fgcode_' + i]) {
            delete metadata[z.title]['fgcode_' + i]
          }

          if (metadata['fgCodeType'][i]) delete metadata['fgCodeType'][i]
          if (metadata[z.title]['Trim_' + i]) {
            delete metadata[z.title]['Trim_' + i]
          }
        } else {
          if (metadata[z.title][i]) delete metadata[z.title][i]
        }
      }

      if (z.datasheet) {
        if (
          proposaldata?.technicaldata?.[z.title] &&
          proposaldata.technicaldata[z.title][i]
        ) {
          delete proposaldata.sizingdata[z.title][i]
        }
      }
    })

    sheet.setRowData(i, rowdata)
  }

  /**
   * @method getColumnName
   * @param {Number} i
   * @return {String}
   * @document column name from index
   */
  getColumnName(i) {
    var letter = ''
    if (i > 701) {
      letter += String.fromCharCode(64 + parseInt(i / 676))
      letter += String.fromCharCode(64 + parseInt((i % 676) / 26))
    } else if (i > 25) {
      letter += String.fromCharCode(64 + parseInt(i / 26))
    }
    letter += String.fromCharCode(65 + (i % 26))

    return letter
  }

  /**
   * @method getColumnNameFromId
   * @param {Array} cellId
   * @return {String}
   * @document column name from index
   */
  getColumnNameFromId(cellId) {
    if (!Array.isArray(cellId)) {
      cellId = cellId.split('-')
    }

    return this.getColumnName(parseInt(cellId[0])) + (parseInt(cellId[1]) + 1)
  }
}
</script>

<style lang="css">
#spreadsheet .jexcel_content table tbody tr td:first-child {
  position: sticky !important;
  left: 0px !important;
}
#spreadsheet .body--light table.jexcel > thead > tr > td {
  border-left: 1px solid #494343;
}

</style>

<style lang="css" scoped>
.jexcel_content {
  height: 100vh;
  width: 100vw;
}
.actinput {
  height: 30px;
  border: 0.5px solid rgba(218, 218, 218, 0.864);
  color: purple;
  text-align: center;
}
</style>
